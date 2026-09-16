# Artifact Tool for Excel: operations, settings and working patterns

**Installed package:** `@oai/artifact-tool 2.8.59` · **Runtime bundle:** `26.909.12148` · **Studied:** 15 September 2026 · **Environment:** Node 24.19.0 on macOS.

## What this reference establishes

Artifact Tool is a JavaScript workbook engine with editable objects, formula calculation, drawing/rendering, Excel import/export, change recording and persistent-session adapters. “Excel” here means the spreadsheet portion of the installed package. It is not automation of the Microsoft Excel application.

The most useful mental model is **read or build → change → calculate → inspect → render → export → reopen and check**. These are separate operations. A call can succeed while calculation returns an incorrect value, a preview can show a chart that Excel export omits, and a saved file can reopen differently from the original in-memory workbook.

This reference covers every spreadsheet operation in the installed help catalogue, every API example in that catalogue, every listed formula's supplied example, the supplied runnable examples, the documented feature families, and the surrounding scripts and workflows. It includes targeted counterexamples and recovery recipes. It does **not** certify every combination of options, every mathematical edge case, native Excel compatibility, or service-dependent Google Sheets behavior.

### Reading routes

1. [Action families](#action-families): what to do, where to do it, and the important inputs and outputs.
2. [Working recipes](#working-recipes): corrected, reusable building blocks.
3. [Calculation findings](#calculation-findings): what the formula tests establish and what they do not.
4. [Inspection and file boundaries](#inspection-and-file-boundaries): understanding existing files, rendering, export and reopening.
5. [Changes, collaboration and sessions](#changes-collaboration-and-sessions): three different ways to wrap mutations.
6. [Suggested workflows and supporting scripts](#suggested-workflows-and-supporting-scripts): the practices around the package.
7. [Limitations and documentation corrections](#limitations-and-documentation-corrections): discrepancies with exact evidence.
8. [Reproduction](#reproduction): rerunnable probes and evidence conventions.
9. [Complete catalogues](#complete-catalogues): all API help records and examples, operation inputs, formula arguments/examples/results, enums, probe receipts and source inventory.

### Evidence conventions

| Label | Meaning |
| --- | --- |
| Documented | A supplied reference, example, help entry or wrapper contract says it exists. |
| Executed | The call completed. This is not a correctness claim. |
| Verified | A stated assertion, independent expected result, serialized-file check or re-import check supports the specific claim. |
| Discrepancy | The observed result differs from the stated expectation or supplied instructions. |
| Conditional / untested | Requires an external service, target application, different platform, or configuration that was not exercised. |

The raw probe harness uses `status: "pass"` to mean **the probe completed without throwing**. Some probes deliberately return a discrepancy as data. Read `detail`, `matches`, export errors and the interpretation here. Counts in the generated catalogues distinguish those cases.

All input data and annotations are synthetic. Existing user workbooks and installed package files were not modified. The captured [source manifest](files/source-manifest.json) records version, paths and SHA-256 hashes. Source snapshots are included so later documentation changes cannot silently change what this study means.

## Action families

### Workbook and worksheet lifecycle

| Actions | Main calls | Inputs and settings | Output and observed behavior |
| --- | --- | --- | --- |
| Create a workbook | `Workbook.create()` | No worksheet names or other arguments | Empty workbook; add sheets explicitly. |
| Load workbook state | `Workbook.load(proto, {validate:true})`, constructor, `toProto()` | Serialized workbook state; optional validation | Editable workbook or serializable object. Protocol state includes sheets, styles, drawings, names, annotations and caches. |
| Add/find/select sheets | `worksheets.add(name \| names[])`, `getItem`, `getItemAt`, `getFirst`, `getActiveWorksheet`, `setActiveWorksheet` | Names and zero-based indexes | Worksheet facade; adding the same name returned the existing worksheet in the test. |
| Count and identify | `getSheetCount`, `getSheetIndex`, `getSheetNameByIndex`, sheet `id`, `sheetId` | Name/index; inspect anchors for reliable identity | Names, indexes, IDs. `id` and `sheetId` are distinct identifiers. |
| Rename, reorder, display | `sheet.name`, `index`, `showGridLines`, `tabColor` | Strings, zero-based index, boolean, color | Mutates worksheet properties. Color getters can return objects. |
| Reset or delete | `sheet.reset(options)`, `sheet.delete()` | `clear: used/none`; `applyTo: contents/formats/all`; delete tables/charts/drawings/sparklines flags | A reset has separate switches for cells and attached objects. |
| Theme | `setColorScheme({name,themeColors})`, `theme` | Theme tokens such as `accent1`, `bg1`, `tx1` with colors | Workbook theme and themed formatting. |

`workbook.sheets` aliases `worksheets`; `definedNames` aliases `names`. A `getItemOrNullObject` result is a facade with `isNullObject`, not a normal JavaScript `null`; testing its truthiness is insufficient.

Evidence: C01–C02, C14, C23; [workbook reference](files/sources/package/api/references/workbook.spec.md), [worksheet reference](files/sources/package/api/references/worksheets.spec.md).

### Ranges, values and formulas

| Actions | Calls/settings | Practical contract |
| --- | --- | --- |
| Address a rectangle | `getRange("A1:C4")`; `getCell(row,col)`; `getRangeByIndexes(row,col,rows,cols)` | A1 text or zero-based coordinates. Utilities such as `columnToLetter(1)` use one-based numbers. |
| Navigate | `getRow`, `getColumn`, `getCell`, `offset/getOffsetRange`, `resize/getResizeRange`, `getResizedRange` | Resize takes the final dimensions; ResizedRange takes dimension changes. |
| Discover occupied areas | `getUsedRange(valuesOnly?)`, `getCurrentRegion()` | Values-only used range differs from a formatting-inclusive range. |
| Read/write contents | `values`, `formulas`, `formulasR1C1`, `write(payload,options)` | Matrices; one-dimensional rows/columns also accepted. Values include strings, numbers, booleans, dates and null. |
| Control writes | `overwrite: allow/error`, `resize: auto/none`, `clear: contents/formats/all` | `overwrite:error` rejected a populated destination. Do not treat declared options as a complete bounds validator. |
| Fill formulas | `fillDown()`, `fillRight()`, `fillFrom(source)` | Relative references move and absolute references stay fixed. FillFrom extends an aligned seed in one direction. |
| Copy | `copyFrom(source,type)`, `copyTo(target,type)` | All, values or formulas. Use copy for a separate destination; fillFrom is not arbitrary paste. |
| Clear/merge | `clear({applyTo})`, `merge(across?)`, `unmerge()` | Contents and formatting can be cleared separately. Drawings have separate deletion methods. |
| Read displayed formulas | `formulasDisplay`, `formulaInfos` | Distinguishes stored formulas from spill/data-table projections, including anchor and editability. |

Important observed details:

- `values = [["=2+3"]]` stores a formula; `"'=2+3"` stores literal text.
- A single-cell target can expand to fit a matrix. A single value assigned to a larger range broadcasts, but a one-cell formula matrix populated only its first cell.
- `range.getRange("A1")` resolved worksheet A1, even when the parent range began at C3. Use relative coordinate methods when that is what you mean.
- Writing a small payload with `clear:"all"` cleared its payload footprint; neighboring cells in the larger selected range remained.
- Dynamic-array projection works, but the tested spill obstruction did not enforce Excel's usual collision behavior. Directly overwriting a projected cell was accepted. An editor should honor `formulaInfos.isEditable`.
- Validate destination bounds and payload dimensions in the wrapper when correctness depends on them.

Evidence: C03–C11, C25, R01; [range contract](files/sources/package/api/references/ranges.spec.md).

### Formatting and layout

Formatting is attached to ranges, with nested objects for font, fill and borders. Assign a compact `range.format` object or edit a nested property.

| Family | Settings |
| --- | --- |
| Font | Name, size, bold, italic, underline and color; see the exact shipped style schemas in the catalogues. |
| Fill and color | Hex strings, theme tokens, RGB/theme color objects, transforms such as opacity/lightening, solid/gradient/pattern configurations. |
| Borders | Individual edges or presets; line style, thickness and color. |
| Alignment | Horizontal/vertical alignment, wrapping; token aliases differ across the facade and serialized schemas. |
| Number display | Excel number-format strings; `setNumberFormat`; format matrices. Number display does not change the stored numeric value. |
| Sizing | `rowHeight` in points, `rowHeightPx`, `columnWidth` in spreadsheet units, `columnWidthPx`; autofit rows/columns. |
| Worksheet view | Gridlines, tab color, freeze rows, freeze columns, unfreeze. |
| Drawing layout | `autoLayoutDrawings(items,{direction,frame,gap,padding})`; vertical, horizontal and grid layouts. |

Thirty pixels of row height read back as 22.5 points. Reading a mixed-range `numberFormat` returned a single format, not a full matrix. Do not infer uniformity from that getter.

Autofit is a starting point. Render and inspect long labels, crowded axes and clipped boundaries. The reviewed diagnostic previews demonstrated actual cell styling, conditional formats and sparklines; the chart data-table preview also showed crowding that still needs an authoring adjustment in a production workbook.

Evidence: C12–C14, F07, A02; [formatting reference](files/sources/package/api/references/formatting.spec.md), [style schema](files/sources/package/references/styles.spec.md), [formatting preview](files/runs/formatting.png).

### Data validation and conditional formatting

**Data validation** defines spreadsheet input rules; it does not validate JavaScript writes automatically. All eight listed types were constructed and exported: `none`, `whole`, `decimal`, `list`, `date`, `time`, `textLength`, `custom`.

Use a range's `dataValidation` setter or the worksheet validation collection. Settings include range, type, comparison operator, formula bounds or list source, blank handling, dropdown, input prompt and error messages/styles. The Office-style list form accepts `{list:{source:[...],inCellDropDown:true},allowBlank:true}`. Writing 999 programmatically into a 1–10 rule succeeded; enforce such bounds yourself before mutation.

**Conditional formatting** attaches rules rather than changing the original cell values. The test exercised 20 rule tokens including aliases: cell comparisons, custom/expression formulas, color scales, data bars, icon sets, text matching, blank/error tests, duplicate/unique values, top/bottom and average comparisons. Rule settings include formulas/operators, thresholds, colors/icons, formatting and rule-specific options. The precise required fields and enum values are preserved in the schemas below.

`workbook.getConditionalFormattingRenderCache(sheetName)` supplies evaluated metadata to renderers. Serialization and representative rendering were checked; accepting an icon-set token does not establish every target application's support.

Evidence: F01–F02; [validation schema](files/sources/package/references/data-validations.spec.md), [conditional-format schema](files/sources/package/references/conditional-formatting.spec.md), [preview](files/runs/conditional-formats.png).

### Three different “table” features

| Feature | Purpose | Main actions | Verified limits |
| --- | --- | --- | --- |
| Excel Table | A named, styled rectangular dataset with structured references | `sheet.tables.add(range,hasHeaders,name)`; get/delete; append rows; name, range, style, headers, totals and banding | Appending with `rows.add(null,rows)` works. Inserting at a specified index throws. Overlapping tables were accepted. |
| What-If Data Table | Recalculate one formula over one/two varying inputs | `sheet.dataTables.add(fullRange,{rowInput,columnInput})`; recalculate; inspect projected results | Use the complete rectangle including corner formula and scenario headers. Local input cells are required. One-variable examples also worked in this version. |
| PivotTable | Aggregate source records by fields | Root `workbook.pivotTables.add(name,source,destination)`; row/column/data hierarchies; layout; cache rebuild | Simple pivot creation/export works, but blank category headings and stale output cells after filtering were observed. |

**Excel Table totals need particular care.** Turning on `showTotals` did not insert a new row. With quantities 2, 3 and 4, a new structured sum changed from 9 to 5 because the final row became the totals row. An already calculated formula still showed 9 until export/re-import, when it also became 5. Reserve the totals row explicitly and test a fresh structured formula. The public `table.columns` and `table.autoFilter` members suggested by broad documentation were undefined here.

**What-If tables:** the method returns `undefined`, not a table handle. New tables need `workbook.recalculate()`. The engine temporarily substitutes the input cells and restores them. A two-input fixture produced exactly the expected 3×3 grid. Both horizontal and vertical one-input fixtures worked and survived re-import, contrary to the supplied skill guidance.

**Pivot slicers:** create the pivot through the root workbook collection when a slicer needs to find it. The tested worksheet-level route could create a pivot but the slicer then reported it missing. Slicer selection and clearing, source edits and cache rebuild were exercised. Filtering to North left an old grand-total cell below the smaller new result, so refresh routines must check the vacated output footprint.

Evidence: C18–C19, F04–F05, R01, R04, R09, A01, A06–A07, V04–V05; [table reference](files/sources/package/api/references/tables.spec.md), [Data Table guide](files/sources/skill/artifact_tool_docs/DATA_TABLES.md).

### Charts

There are **25 chart-type tokens** in the installed help enum. Every type was constructed and exported; rendering completed for 24. The box-and-whisker render stalled and its process was stopped. The export audit found native chart XML for 16 types; nine were omitted by the exporter. The complete per-type results appear in the generated catalogue.

| Actions/settings | Inputs and behavior |
| --- | --- |
| Create | `sheet.charts.add(type,range,seriesBy?)` or a configuration object with categories/series; compatibility names such as ColumnClustered are also documented. |
| Bind data | Worksheet ranges, nonadjacent addresses, row/column orientation, category formulas, value formulas or explicit arrays. |
| Rebind | `chart.setData(...)` updates sources; it rebuilds series, so apply per-series style afterwards. |
| Position | `setPosition(startCell,endCell)`, range anchors, zero-based `from`, `to` or pixel `extent`; width and height. |
| Titles and legend | Title text/text style, legend visibility, position and text style. |
| Axes | Category/value or x/y facade; number formats, bounds, units, placement, gridlines, titles, label formatting. |
| Series | Names, values/categories, fill/line, markers, data labels, trendlines and error bars. |
| Chart-specific | Bar direction/grouping/gap/overlap; pie/doughnut and scatter options; see exact API notes and schemas. |
| Analytics | Trendline types and options, including moving average; x/y error bars and value/type settings. |
| Data labels/table | Show values/category/series, position, number format, connector/text settings; chart data table and legend keys. |
| Convert | `await workbook.chartToImage(sheetIndex,chart.id)` returns bytes, content type and pixel size. It did not replace/delete the source chart. |
| Remove | Chart `delete()`; collection `deleteAll()/clear()`. |

**Corrected mutation pattern:** use `chart.dataLabels.showValue = true` and `chart.dataTable.visible = true`. Four help examples assign a plain object to these properties. Those calls complete, but later export throws `toProto is not a function`. Creation-time data-label configuration worked; creation-time data-table visibility was ignored in the test.

For a range-bound series, `series.values` and `resolveValues()` returned empty arrays while the source formula, rendered plot and exported chart cache showed the updated value 200. Inspect the worksheet source, formula references and actual export, rather than relying on that getter alone.

Evidence: CH-* probes, F08, R07–R08, A02, API examples 2/20/21/22; [chart settings preview](files/runs/chart-settings.png), [chart guidance](files/sources/skill/features/charts.md).

### Images, shapes and sparklines

| Family | Actions and inputs | Outputs and limits |
| --- | --- | --- |
| Images | Add SVG text, bytes plus contentType, data URL or file path; replace content; remove; anchor at creation | Real image content was rendered/exported. `uri` and `prompt` fields do not themselves prove image fetching/generation. |
| Image paste | `workbook.fromImage(sheetIndex,{bytes,contentType},target)` | Inserts an image, not OCR-derived cells. The test used the first target cell and original pixel dimensions rather than fitting the full supplied rectangle. |
| Shapes | Add geometry, anchor, fill and line; set text; position/layout; delete | All 190 enum tokens were attempted. 189 built-in geometries were created, exported and included in rendered contact sheets. The remaining `custom` geometry also rendered, exported and re-imported after supplying its required path data. |
| Sparklines | `sparklineGroups.add({targetRange,sourceData,type,...})`; getAll/getGroupForCell; delete/clear | Line, column and stacked/win-loss types were rendered and exported. Range-based convenience API and deprecated alias were also exercised. |
| Sparkline settings | High/low/first/last/negative markers; line weight; colors; axes/scale; empty/hidden-cell handling | Manual bounds alone do not select custom scaling. Date-axis configuration has documented support limits. |

Do not replace the image's anchor facade with a plain object after creation: that broke serialization. Set the anchor during creation or use the verified `image.set` operation. Sparkline groups do not have the assumed `group.delete()`; use `sheet.sparklineGroups.delete(group)`.

Custom geometry takes `customPaths`, each with positive pixel `width`/`height` and `commands`. The verified example uses `moveTo`, `lineTo`, `quadBezTo`, `cubicBezTo` and `close`. Coordinates must be finite. This configuration was established through a [targeted installed-source excerpt](files/custom-path-source.json) and the [executed workflow example](workflow-examples.mjs), because the supplied spreadsheet docs do not describe the required paths.

Evidence: F03, F06–F07, R05–R06, [geometry results](files/runs/geometry.json), [images](files/runs/images.xlsx), [sparklines preview](files/runs/sparklines.png), [custom-shape and scenario receipt](files/runs/workflow-example.json).

### Defined names, comments and notes

| Family | Actions/settings | Observations |
| --- | --- | --- |
| Named ranges | Workbook/sheet `names.addRange(name,address,{description})`; lookup/delete | Can be used in formulas; workbook and sheet scope were exercised. |
| Named functions | `names.addFunction(name,{lambda,description,parameters,returns})` | The property must be `lambda`. A conflicting example used `formula` and produced #NAME?. |
| Threaded comments | Set author; add thread on cell/range; reply; resolve/reopen; react | Threads, replies and authors were exported as native parts. Direct reply reactions worked. |
| Legacy notes | Low-level `workbook.notes.add(config)`; delete | Native cell note export worked with the complete target/author/body configuration shown below. |
| Presence/awareness | Selection/presence objects; collaborative editing metadata | Surface exists, but the tested workbook.apply presence/reaction commands only returned TODO warnings. |

The note operation `note.add` is not implemented by the tested `workbook.apply` dispatcher. This differs from the working notes collection. Do not assume similar names across dispatchers have identical support.

Evidence: C15–C17, R02, R10–R11, A03–A04, V02–V03; [names/annotations reference](files/sources/package/api/references/comments-notes-names.spec.md).

## Working recipes

These fragments use existing `workbook` and `sheet` variables unless shown otherwise. They are small versions of the executed probes, with version-specific corrections incorporated.

### 1. A workbook with a real formula dependency

```js
import { Workbook, SpreadsheetFile, FileBlob } from "@oai/artifact-tool";

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Data");
sheet.getRange("A1:C3").values = [
  ["Quantity", "Price", "Revenue"],
  [2, 10, null],
  [3, 10, null],
];
sheet.getRange("C2").formulas = [["=A2*B2"]];
sheet.getRange("C2:C3").fillDown();
workbook.recalculate();
const results = sheet.getRange("C2:C3").values; // [[20],[30]]

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save("model.xlsx");
const reopened = await SpreadsheetFile.importXlsx(
  await FileBlob.load("model.xlsx")
);
```

Write rectangular data in one operation, then formulas in blocks or by seed-and-fill. Keep the builder as the source of truth for new workbooks; inspect and make targeted changes for existing workbooks.

### 2. Named function

```js
workbook.names.addFunction("AddTax", {
  lambda: "LAMBDA(amount,amount*1.1)",
  description: "Apply a synthetic ten percent tax rate",
  parameters: [{ name: "amount", description: "Amount before tax" }],
  returns: "Amount including tax",
});
sheet.getRange("B1").formulas = [["=AddTax(100)"]];
// Approximately 110, subject to normal floating-point precision.
```

The quick start recommends Excel's `_xlfn.` function and `_xlpm.` parameter prefixes for relevant modern formulas at export. Both plain and prefixed examples calculated here. Native Excel behavior was not tested, so keep the destination-specific guidance visible.

### 3. A two-variable What-If Data Table

```js
sheet.getRange("B2:B3").values = [[2], [3]];
sheet.getRange("B5").formulas = [["=(B2+B3)*10"]];
sheet.getRange("E4").formulas = [["=B5"]];
sheet.getRange("F4:H4").values = [[1, 2, 3]];
sheet.getRange("E5:E7").values = [[1], [2], [3]];
sheet.dataTables.add("E4:H7", { rowInput: "B2", columnInput: "B3" });
workbook.recalculate();
// F5:H7 = [[20,30,40],[30,40,50],[40,50,60]]
```

The top-left formula, row scenarios and column scenarios are part of the argument. Passing only F5:H7 fails. The inputs must be the cells actually driving the model, not unconnected local copies.

### 4. Safe chart facade mutation

```js
const chart = sheet.charts.add("line", sheet.getRange("A1:C3"));
chart.setPosition("F2", "N18");
chart.title.text = "Synthetic series";
chart.dataLabels.showValue = true;
chart.dataTable.visible = true;
chart.dataTable.showLegendKey = true;
const picture = await workbook.chartToImage(0, chart.id);
// picture: { bytes, contentType, widthPx, heightPx }
```

A chart image is an additional output. Explicitly add it as a worksheet image only if that is the requested workflow.

### 5. A native cell note

```js
workbook.comments.setSelf({ displayName: "Example author" });
const note = workbook.notes.add({
  id: "example-note",
  target: {
    cell: { sheetName: sheet.name, sheetId: sheet.sheetId, address: "A1" }
  },
  authorId: workbook.comments.self.id,
  createdAt: new Date().toISOString(),
  body: { plainText: "Synthetic source note" },
});
```

This is a version-specific collection contract, established through the installed facade and successful native note export. Prefer the documented threaded-comments API for review discussions.

## Calculation findings

### What was actually tested

The help catalogue lists **494 formula records across 12 categories**. Every supplied example was evaluated in a separate workbook fixture, with Sheet1 and Data holding a 20×5 numeric grid. Results include the formula, top-left value, projected 4×4 output and elapsed time. Each worker had a hard timeout.

That sweep is a **breadth test**, not an independent correctness test. Some examples reference context not present in the common fixture. A #VALUE! or #NUM! from such an example cannot by itself establish that the function is unsupported. A returned number cannot by itself establish that it is correct.

The broad run returned 45 explicit “not implemented” outputs, 90 Excel-style error outputs and 33 exception-like text outputs. The complete input, arguments and observed output for every function are in the formula catalogue.

A separate set of **77 known-result probes matched 72 expectations**. Its five discrepancies were OFFSET, INDIRECT, FORMULATEXT and ISFORMULA returning unimplemented text, and COUNTIF returning zero for five blank cells. The [oracle receipt](files/runs/oracles.json) records each expected and actual result.

A further fixture tested all 12 database functions with a real header row and criteria range. Only two results matched the expected answers, and one of those was a coincidental minimum. For example, DSUM returned 0 instead of 40; DCOUNT returned 4 instead of 2; DAVERAGE returned 25 instead of 20. These results do not support using the database family for criteria-sensitive work without independent verification.

### Consequences for wrappers

1. **Check expected outputs, not only error tokens.** Unimplemented functions can return ordinary text, so a scan for #REF!/#DIV/0!/#VALUE! misses failures.
2. **Check prerequisites and missing data explicitly.** COUNTIF/COUNTIFS blank behavior differed from COUNTBLANK. Valid zero, empty cell and missing source are different cases.
3. **Use bounded matching ranges.** A whole-column SUMIFS fixture returned #VALUE! where the equivalent bounded formula returned 3. ROWS(A:A) returned the populated extent, 4, rather than Excel's full row count.
4. **Test input changes.** Change a source driver, recalculate, inspect the dependent value, export and reopen. The table-totals example showed why the first cached value is insufficient.
5. **Trace dependencies when a result is surprising.** `workbook.trace("Sheet!A1")` returns formula/value/precedent trees. `collectFormulaUsageStats()` can be empty for formulas using only arithmetic operators.
6. **Separate target-engine support.** The skill recommends OFFSET/INDIRECT patterns, but this installed calculator did not implement them. Their appearance in an XLSX formula does not prove local calculated values are usable.

Evidence: C24–C27, all FX-* oracle probes, A06, V04; [formula results](files/runs/formula-results.json), [formula references](files/sources/package/api/references/formulas.spec.md).

## Inspection and file boundaries

### Discovery, inspection, search and resolution

| Tool | Main settings | Return value and purpose |
| --- | --- | --- |
| `workbook.help(query,options)` | Query such as `worksheet.charts.add`, `fx.*`, `enum.ChartType`; search; include index/examples/notes; maxChars | Bounded NDJSON help, record count and truncation. |
| `workbook.inspect(options)` | Kind, target ID and before/after lines, include/exclude, search, maxChars; range in supported examples | Records/NDJSON plus metadata and notices. Understand sheets, tables, formulas, charts, threads and names. |
| `workbook.findCells(...)` | Search term, sheetId, match formulas, max results and supported search options | Matches with address, value/formula metadata, total and truncation. |
| `workbook.resolve(anchor)` | Actual `wb/`, `ws/`, `ch/`, `th/` anchors from inspection | Editable facade. Placeholder IDs in docs must be replaced. |
| `workbook.trace(cell)` | Qualified cell address | Formula and precedent tree; missing target returned null. |
| Layout export | Sheet/range selection and crop | Geometry, styles, dimensions, merges, validation, names and drawing anchors for template analysis. |

Use small targeted inspection for ordinary editing. The exhaustive help queries in this study deliberately inventory the entire package; they are not the normal recommended authoring pattern. Structured help examples require mapping their fields into actual method arguments. One help example had a malformed trailing dot and wrong call shape; its repair is explicitly recorded.

### Input routes

| Route | Input | Output and caveats |
| --- | --- | --- |
| Native Excel | `FileBlob.load(path)` → `SpreadsheetFile.importXlsx(blob)` | Editable workbook. Feature preservation needs checks before and after re-export. |
| CSV | `Workbook.fromCSV(text,{sheetName})` | Workbook; quotes, embedded commas and newlines handled in the fixture. Numbers remained strings until converted. |
| Instance CSV | `workbook.fromCSV(...)` | Documented `{sheet,range}` result, but behavior depended on collaborative state. A populated freshly created workbook threw a hydration error. Prefer static import unless this state transition is tested. |
| Markdown | `Workbook.fromMarkdown(table,{sheetName,format})` | Workbook; formatted mode creates a native table/autofits. Numeric-looking values remained strings in the test. |
| HTML | `toHTML(sheetIndex,range,{formulas})`; `fromHTML(sheetIndex,html,target)` | Copy/paste data and optional formulas. Formulas were preserved literally, not relocated to the new destination. |
| Image | `fromImage(...)`, worksheet images | Drawing insertion; not OCR. |
| Serialized state | `toProto()/Workbook.load(proto)` | Local state cloning/checkpoint; distinct from a shared CRDT history. |
| Google Sheets | `fromGoogleSheets(config)`, `configureGoogleSheets(config)` | Documented adapter, not exercised against an account in this study. |

### Output routes

`SpreadsheetFile.exportXlsx(workbook)` returns an object with `.save(path)`. Outside a managed session, `workbook.render` and `workbook.export` return Blob-like objects that can be read with `arrayBuffer()` or `text()`.

`workbook.export` supports `png`, `jpeg`, `layout` and `xlsx`. Raster options include sheet/name/index, range, center, width/height, scale, crop, headers and JPEG quality. Layout supports sheet/range/crop; raster framing options are not interchangeable with layout semantics.

The managed session wrapper changes these return contracts: it writes files and returns **saved-file descriptors**. Do not call Blob methods on those descriptors.

Rendering can also keep the Node process alive after the requested work finishes. In the bounded F09 retry, all four output variants were written and the receipt completed in 433 ms, but the process remained open until the timeout. Distinguish operation completion from process shutdown when designing a command-line wrapper.

### Verification has four distinct layers

1. **Values:** compare exact or tolerance-based expected results and formula strings.
2. **Rendering:** look at the actual image for clipping, spacing, chart readability and conditional-format behavior.
3. **Native package:** inspect the XLSX ZIP's XML for formulas, native charts/tables, validation, sparklines, comments, pivots/slicers and caches. The audit script is read-only; it does not repair exports.
4. **Reopening:** re-import and compare important values and feature structures, then export again if that is the intended workflow.

The included XML audit verifies parsing and the presence/content of selected parts. This is not full Open XML schema validation and does not prove that Microsoft Excel will accept, display or recalculate every feature identically.

Fifteen selected files were also imported, recalculated and re-exported. The tested native comment, validation, conditional-format, sparkline, chart, table and name part counts remained stable. Two specific losses appeared: slicer XML disappeared while its cache remained, and original SVG media disappeared while raster image content remained. The complete before/after part comparison is in the catalogue. A successful first export therefore does not establish safe repeated editing of every feature.

## Changes, collaboration and sessions

### 1. Direct mutable objects

The normal builder imports the package, owns an in-memory workbook, and calls its facades. Ordinary JavaScript can prepare data, loop across periods, construct matrices, assert expected results and save files.

For a small reversible experiment, serialize an initial workbook, operate on a clone and keep the original unchanged. For scenario capture, change the actual input cells, calculate, capture labeled outputs and restore inputs in a `finally` block. Captured results are snapshots with an explicit refresh method; they are not simultaneously live scenarios.

### 2. Recorded operations and collaborative updates

```js
const { result, patch, idMap, crdtUpdateV2 } = workbook.record(() => {
  sheet.getRange("A1").values = [[2]];
  sheet.getRange("B1").formulas = [["=A1*3"]];
  return "changed";
});
const applied = anotherWorkbook.apply(patch);
if (applied.warnings.length) throw new Error(applied.warnings.join("\n"));
```

The operation catalogue contains **39 commands and 42 examples**. All documented examples executed without dispatcher warnings after fixture preparation. Some were legitimate no-ops, such as adding a sheet that the fixture already contained. The catalogue records the actual input, warnings and whether serialized state changed.

Important distinctions:

- `record/recordAsync` returns the callback result, a readable patch, ID map and binary CRDT update.
- `apply` accepts operations and returns warnings/ID mapping. Unknown operations can warn without throwing.
- A valid operation followed by an unknown operation still changed the workbook. Do not assume atomic rollback for a patch batch.
- Undo/redo was tested after `recordAsync`: 1 → 2 → 1 → 2.
- CRDT updates require a **shared initialization history**. Two independent loads of identical serialized state did not form a valid shared replica. Capturing the initial hydration update, applying it to an empty replica, then applying the edit update successfully produced 99.
- `onCrdtUpdateV2` returns an unsubscribe function. `applyCrdtUpdateV2(update,{recalculate:true})` applies a remote update. The study verifies a simple replica transfer, not all concurrent merge cases.

Evidence: C28–C29, R11, V01, V03; [operation receipts](files/runs/ops.json).

### 3. Persistent artifact sessions

The package also ships a local MCP server, exposed as `artifact-session-mcp` / `@oai/artifact-tool/session-mcp`. It has one tool, `artifact_session_run`, which executes an asynchronous JavaScript function body with prebound `artifact`, `session`, `input` and `workbook` for spreadsheets.

| Setting | Contract |
| --- | --- |
| `code` | Required asynchronous JS body; return a JSON value for the next call. |
| `artifactType` | Spreadsheet or presentation; spreadsheet default. Only spreadsheet behavior was studied. |
| `target` | Existing Office-file path, snapshot path or stable session ID. Existing XLSX target imports that file. |
| `input` / `summary` | JSON input for code; summary stored with a committed change. |
| `create` / `title` | Create missing targets; initial title. |
| `sessionRoot` / `snapshotPath` | Persistent-session storage location. |
| `sourcePath` | Persistence metadata override. It does **not** import a file. |
| `outputDirectory` | Parent for per-run saved renders, inspections and exports. |
| `timeoutMs` | 1–300,000 ms; default 120,000. |
| `sessionMode` | With no explicit target, reuse the current task's session or create a new one. Requires task context for the implicit route. |

The exact JSON schema is [saved here](files/runs/session-mcp-schema.json). Unknown fields were rejected. The source also imposes limits on code/input/result payload sizes; this is a compact-result interface, not a route for returning whole workbook buffers.

**Executed session sequence:** create total 12 → reopen in a different worker → change quantity and get total 20 → save inspection/PNG/XLSX descriptors → deliberately throw after a mutation → reopen and verify the prior value remained → import an existing XLSX → test missing target/unknown argument → test task-default session metadata → time out a worker and reopen.

Session output includes success/commit state, IDs, state version, snapshot path, output descriptors, timing/worker metadata and callback result. A successful read-only/render call can return `committed:false`; that does not mean the call failed. A thrown callback returned `not_committed` and preserved saved state. A forced timeout reported an unknown execution outcome; reopen and inspect before retrying a mutation.

The separate first-party `artifactSession.run/status/abort` client depends on its host REPL/RPC environment. Calling it in ordinary Node produced its explicit host-required error. The local MCP route was exercised directly; first-party host integration and cancellation UI were not imitated.

Evidence: [session probe script](sessions.mjs), [session receipts](files/runs/sessions.json), V06.

## Suggested workflows and supporting scripts

The package's methods, the platform wrapper and the Spreadsheets skill serve different roles. A method defines an operation; a wrapper manages execution/state/files; the skill gives workbook design and verification guidance. Guidance can be useful while still overstating a particular engine's support.

### The documented authoring workflow

1. Establish the audience, decision, inputs, outputs, formulas, target engine and required checks.
2. Inspect/render a supplied workbook before editing; preserve its structure and conventions.
3. Load the bundled runtime. Create a local dependency symlink/junction in the working directory. Keep the installed dependencies unchanged.
4. Run the supplied operation marker once before first authoring.
5. Maintain one rerunnable `.mjs` builder. Prepare rectangular data and write in blocks. Calculate reusable results once and link to them.
6. Recalculate; run meaningful independent checks; test a representative input/scenario change.
7. Export and visually inspect relevant sheets/ranges. Correct readability problems in the builder.
8. Save the native file; reopen and check load-bearing formulas and features. Deliver the requested artifact and explain material limitations.

For this exploration, separate probe scripts intentionally isolate potentially failing features and preserve initial failures. That is an experimental harness, not a recommendation to scatter a normal workbook across many competing builders.

### Workbook-design paradigms in the supplied skill

| Pattern | What it asks the author to do | How it relates to the package |
| --- | --- | --- |
| Inputs → assumptions → builds → outputs | Preserve source data; calculate meaningful steps once; link outputs to them | Formula/range operations implement it. Roles can share one sheet. |
| A single active case | One authoritative selector, grouped active assumptions, one set of schedules | Use verified CHOOSE/lookup patterns. OFFSET is suggested but locally unimplemented. |
| Meaningful missing-data handling | Distinguish valid zero, blank, unavailable and not applicable | Add explicit source checks; validation rules do not guard JS writes. |
| Independent terminal checks | Reconcile against independent evidence; keep check results out of business logic | Assertions or worksheet check areas; no separate Checks tab needed for a trivial workbook. |
| Narrow edits | Match nearby formulas/styles and preserve unrelated features | Inspect, resolve, write only relevant cells, then verify affected dependencies. |
| Native features where requested | Preserve Excel Tables, pivots, What-If tables, charts and formulas | Test native export; a rendered imitation does not establish native preservation. |
| Scenario capture | When explicitly required, capture case results and show refresh/staleness | JavaScript change/calculate/capture/restore is feasible. Circular self-retaining formulas were not certified. |
| Template reconstruction | Inspect values/formulas plus rendered/layout evidence; author semantic code | Layout JSON is evidence, not a recommended opaque replay format. |
| Presentation by audience | Clear main answer; useful units/periods; limited tabs; readable source and calculations | Formatting and drawing APIs implement presentation, not the analytical judgment. |

### Supporting files and what they actually do

| Supplied file | Role and execution status |
| --- | --- |
| [Operation marker](files/sources/skill/container_tools/mark_artifact_operation_started.mjs) | Validates operation kind, expected output count and format. A valid invocation exits silently; the script itself emits no marker text and writes no file. The skill requires the invocation before authoring. It does not build, calculate or export a workbook. |
| [Quick-start example](files/sources/package/examples/quick_start_example.ts) | Runnable basic authoring pattern. Executed with the bundled Node runtime. |
| [Chart suggestions](files/sources/package/examples/chart_suggestions.ts) | Runnable chart usage examples. Executed. |
| [Formula trace/help example](files/sources/package/examples/formula_trace_and_help.ts) | Demonstrates introspection and formula diagnostics. Executed. |
| [Existing-workbook inspection example](files/sources/package/examples/inspect_existing_workbooks.ts) | Demonstrates reading an imported workbook. Executed after replacing its source-tree-relative import with the installed package import. |
| [Create/edit workflows](files/sources/skill/workflows/create_workflows.md) | Workbook structure, checks and targeted editing guidance; not additional executable APIs. |
| [Style guidance](files/sources/skill/style_guidelines.md) | Visual hierarchy, number/date display, spacing, readable charts and user/reference precedence. |
| [Finance guidance](files/sources/skill/domain_guidance/financial_models.md) | Periods, accounting bases, assumptions, schedules, cash/valuation and independent reconciliations; finance color/number conventions. |
| [Marketing guidance](files/sources/skill/domain_guidance/marketing_advertising.md) | Matching populations, periods and denominators; source-based performance measures and appropriate aggregation. |
| [Healthcare guidance](files/sources/skill/domain_guidance/healthcare.md) | Administrative/clinical context, units, denominators, source limits and traceable calculations. |
| [Scientific guidance](files/sources/skill/domain_guidance/scientific_research.md) | Preserve original observations, units, transformations and uncertainty; use appropriate independent checks. |
| [Read-only Q&A](files/sources/skill/references/read_only_qna.md) | Inspect and answer without authoring merely to answer a question. |
| [Image references](files/sources/skill/references/image-references.md) and [template elicitation](files/sources/skill/references/template-elicitation.md) | Obtain sufficient source/layout information before reconstruction; images are evidence with limitations. |
| [Google Sheets routing](files/sources/skill/routing/google_sheets.md) | Destination-specific creation/editing path. Requires connected Google services; not exercised against an account here. |
| [Agent metadata](files/sources/skill/agents/openai.yaml) | Skill discovery metadata, not spreadsheet execution behavior. |

All four runnable package examples completed. Their adaptations, logs and artifacts are recorded in [the example receipt](files/runs/supplied-examples.json).

The operation-marker script accepts exactly six positional arguments in this order: `--operation-kind create|edit --expected-output-count 1–100 --output-format csv|tsv|xls|xlsm|xlsx`. The count must be a safe integer; the format is case-insensitive. Invalid arguments print usage to stderr and set exit code 2. These accepted format tokens are bookkeeping, not proof that the workbook exporter supports all five formats.

### Useful code around the tools

The finite, reusable wrapper jobs are:

- **Data preparation:** parse/coerce dates and numbers, preserve raw data, join by explicit keys, check duplicates and matching ranges, prepare block writes.
- **Builder orchestration:** one script, explicit inputs/outputs, correct dependency/runtime, clear output directory and deterministic worksheet references.
- **Verification:** exact/tolerance-based expected values, boundary cases, source controls, formula preservation, changed-input checks, error/text scans and native-file inspection.
- **Scenario execution:** mutate authoritative inputs, recalculate, capture labeled results, restore in `finally`, record the refresh method.
- **Inspection-driven editing:** bounded inventory → resolve actual IDs → targeted mutation → focused after-inspection.
- **Template work:** render plus layout/values/formulas → semantic reconstruction → visual comparison.
- **Recorded edits:** inspect warnings, retain a checkpoint, replay patches against the intended starting state.
- **Persistent execution:** stable target/session ID, JSON input/result, saved output descriptors, reopen after ambiguous timeouts.
- **Failure isolation:** worker process per risky example with timeout; retain original error and separately record the corrected attempt.
- **Export assurance:** preserve warnings/logs, examine native ZIP parts, reopen, and distinguish native features from images or static values.

These do not require a new package API for every loop or check. They are ordinary code composed around the operations documented here.

## Limitations and documentation corrections

| Issue | Evidence | Working interpretation or correction |
| --- | --- | --- |
| Plain chart dataLabels/dataTable assignment breaks export | API 2/20/21/22; A02 | Mutate existing nested facades. |
| Image anchor object replacement breaks export | F06; R06 | Set at creation or use the tested image.set operation. |
| Named-function example uses wrong property | C16; R02 | Use `lambda`, not `formula`. |
| Data-table body-only rectangle is wrong | F05 | Supply the full rectangle with corner formula and scenarios. |
| One-variable tables declared unsupported | R04 | Both orientations worked and re-imported in 2.8.59. This is version-specific evidence. |
| OFFSET/INDIRECT guidance exceeds local calculator support | FX oracles | Choose a verified equivalent where within scope; do not trust local cached results. |
| Blank COUNTIF and full-column calculations differ | C27; FX-COUNTIF-blank | Use independently checked missing-data logic and bounded ranges. |
| Database criteria calculations disagree | A06 | Avoid relying on that family without independent expected-result tests. |
| Enabling table totals changes data interpretation and leaves stale cache | V04 | Reserve totals row; inspect new formulas and reopened file. |
| Table column/filter editing surface absent | R01 | Do not invent these methods from broad prose descriptions. |
| Unsupported chart types disappear during native export | Per-type native audit | Consult the full chart matrix. Render success alone is insufficient. |
| boxWhisker render stalls | Feature run stopped after approximately 90 seconds | Isolate this type with a hard process timeout. |
| Slicer lookup depends on pivot creation route | R09; A01 | Use root workbook pivot creation in this version. |
| Filtered pivot leaves old output cells | V05 | Verify/clear vacated output area through a tested refresh process. |
| Dispatcher can warn and still retain earlier mutations | V03 | Inspect warnings; use checkpoints when all-or-nothing behavior is needed. |
| CRDT replicas need shared history | C29; V01 | Initialize from the same CRDT baseline update. |
| Instance CSV import has state restrictions | R03 | Prefer static import and an explicit copy/merge plan. |
| HTML formula paste does not relocate references | C22 | Use formula-aware copy operations or explicitly transform references. |
| Custom shape requires path configuration | Geometry sweep; workflow example | Supply `customPaths` with dimensions and commands. Corrected example rendered, exported and re-imported. |
| Slicers and SVG originals lost on re-export | Selected feature roundtrips | Slicer XML vanished while its cache remained; SVG originals were replaced by raster-only media. Check the feature preservation needed by the actual workflow. |
| First-party sessions need their host | V06 and sessions | Use the supported host or the separately tested local MCP route. |

The supplied edit guidance asks authors to preserve protection, hidden/grouped rows, external links, shared formulas and calculation/iteration settings. That instruction is **not** a verified public authoring contract for all those features. This study does not claim VBA/macro authoring, live Excel automation, Excel application settings, certified iterative calculation, OCR, external-data refresh or complete Google Sheets parity.

## Reproduction

The scripts load the installed package through the local `node_modules` link created by [prepare.mjs](prepare.mjs). For another computer, set `ARTIFACT_DEPENDENCIES` to the bundled Node modules directory and `SPREADSHEET_SKILL` to the installed Spreadsheets skill. Use the bundled Node and Python executables returned by the workspace-dependency loader.

Run from this folder, with `node` and `python` below referring to those bundled executables:

```sh
node prepare.mjs
node explore.mjs discover
node explore.mjs core
node chart-sweep.mjs
node explore.mjs ops
node formula-sweep.mjs
node explore.mjs oracles
node api-sweep.mjs
node explore.mjs followup
node explore.mjs advanced
node explore.mjs recovery
node run-examples.mjs
node sessions.mjs
node final-coverage.mjs
node workflow-examples.mjs
python export-audit.py
node build-reference.mjs
```

The initial feature run encountered the boxWhisker stall. Use the bounded chart runner below for a complete rerun instead of launching the unbounded all-feature mode. The original interrupted receipts and corrective follow-ups are intentionally retained. A rerun can replace generated receipts, so copy the study folder first if preserving this exact dated run.

The [main probe driver](explore.mjs) also accepts `PROBE_FILTER` (a regular expression over probe IDs) and `RUN_LABEL` (a separate receipt suffix). Formula/API sweeps use separate worker processes and hard timeouts. [final-coverage.mjs](final-coverage.mjs) isolates the geometry and selected-feature re-import runs with 45-second limits.

### How to interpret the saved evidence

- [Source manifest](files/source-manifest.json): package metadata, exact source paths and hashes.
- [API help](files/runs/api-help.ndjson): original catalogue; not all claims are reliable.
- [API execution results](files/runs/api-example-results.json): all examples with corrections/prerequisites and export outcomes.
- [Formula results](files/runs/formula-results.json): common-fixture breadth run.
- [Known-result checks](files/runs/oracles.json): independent expected values.
- [Export audit](files/runs/export-audit.json): native ZIP contents, formulas and selected feature structures.
- [Session receipts](files/runs/sessions.json): exact inputs and returned state/output contracts.
- [Probe driver](explore.mjs): actual fixtures and assertions. Probe IDs in this document point to these cases.
- The generated catalogues below place the finite action inventory, settings, inputs and observations in this same Markdown document.

<!-- GENERATED CATALOGUES -->

# Complete catalogues

These appendices preserve the finite installed catalogue. “Executed” means the call ran in its stated fixture; consult the findings above for correctness and export limits. Each code sample is the supplied input unless an adaptation is explicitly recorded.

## Coverage receipt

```json
{
  "packageVersion": "2.8.59",
  "sourceDocuments": 73,
  "apiRecords": 92,
  "apiExamples": 126,
  "apiExecuted": 126,
  "apiExported": 122,
  "operations": 39,
  "operationExamples": 42,
  "formulaRecords": 494,
  "formulaResults": {
    "value returned": 326,
    "unimplemented text": 45,
    "Excel error": 90,
    "exception-like text": 33
  },
  "independentFormulaChecks": 77,
  "independentFormulaMatches": 72,
  "chartTypes": 25,
  "nativeChartTypes": 16,
  "shapeTypes": 190,
  "builtInShapeCreations": 189,
  "customShape": {
    "captures": [
      {
        "price": 8,
        "values": [
          [
            16
          ],
          [
            24
          ]
        ]
      },
      {
        "price": 10,
        "values": [
          [
            20
          ],
          [
            30
          ]
        ]
      },
      {
        "price": 12,
        "values": [
          [
            24
          ],
          [
            36
          ]
        ]
      }
    ],
    "restored": [
      [
        10
      ],
      [
        10
      ]
    ],
    "independentClone": true,
    "errorScan": [],
    "customShapeId": "h2orqa",
    "reimportedShapes": 1,
    "reimportedValues": [
      [
        20
      ],
      [
        30
      ]
    ]
  },
  "auditedXlsxFiles": 195,
  "xmlParseFailures": 0,
  "selectedRoundtrips": 15
}
```
## Chart-type matrix

Every enum token was tested with the same simple numeric fixture. Rendering is an engine-output check, not a judgment that the fixture is economically suitable for every chart type.

| Type | Native chart XML | Render | Export evidence |
| --- | --- | --- | --- |
| `line` | lineChart | Completed | [XLSX](files/runs/chart-line.xlsx) |
| `pie` | pieChart | Completed | [XLSX](files/runs/chart-pie.xlsx) |
| `bar` | barChart | Completed | [XLSX](files/runs/chart-bar.xlsx) |
| `doughnut` | doughnutChart | Completed | [XLSX](files/runs/chart-doughnut.xlsx) |
| `scatter` | scatterChart | Completed | [XLSX](files/runs/chart-scatter.xlsx) |
| `bubble` | bubbleChart | Completed | [XLSX](files/runs/chart-bubble.xlsx) |
| `radar` | radarChart | Completed | [XLSX](files/runs/chart-radar.xlsx) |
| `treemap` | Omitted | Completed | [XLSX](files/runs/chart-treemap.xlsx) |
| `sunburst` | Omitted | Completed | [XLSX](files/runs/chart-sunburst.xlsx) |
| `map` | Omitted | Completed | [XLSX](files/runs/chart-map.xlsx) |
| `waterfall` | Omitted | Completed | [XLSX](files/runs/chart-waterfall.xlsx) |
| `line3D` | line3DChart | Completed | [XLSX](files/runs/chart-line3D.xlsx) |
| `pie3D` | pie3DChart | Completed | [XLSX](files/runs/chart-pie3D.xlsx) |
| `area3D` | area3DChart | Completed | [XLSX](files/runs/chart-area3D.xlsx) |
| `bar3D` | bar3DChart | Completed | [XLSX](files/runs/chart-bar3D.xlsx) |
| `funnel` | Omitted | Completed | [XLSX](files/runs/chart-funnel.xlsx) |
| `histogram` | Omitted | Completed | [XLSX](files/runs/chart-histogram.xlsx) |
| `boxWhisker` | Omitted | Timed out (10-second bounded retry) | [XLSX](files/runs/chart-boxWhisker.xlsx) |
| `stock` | stockChart | Completed | [XLSX](files/runs/chart-stock.xlsx) |
| `surface3D` | surface3DChart | Completed | [XLSX](files/runs/chart-surface3D.xlsx) |
| `ofPie` | ofPieChart | Completed | [XLSX](files/runs/chart-ofPie.xlsx) |
| `surface` | surfaceChart | Completed | [XLSX](files/runs/chart-surface.xlsx) |
| `pareto` | Omitted | Completed | [XLSX](files/runs/chart-pareto.xlsx) |
| `combo` | Omitted | Completed | [XLSX](files/runs/chart-combo.xlsx) |
| `area` | areaChart | Completed | [XLSX](files/runs/chart-area.xlsx) |

## Selected export/re-import comparisons

All 15 selected files imported, recalculated and re-exported after correcting the harness to use `sparklineGroups.getAll()`. Counts below describe native parts, not full semantic equivalence. SVG source parts disappeared while image facades remained; slicer parts disappeared while caches remained.

| Fixture | Important parts before → after | Lost ZIP parts |
| --- | --- | --- |
| comments | comments: 3 → 3 | None |
| native-note | comments: 2 → 2 | None |
| sparklines | sparklines: 3 → 3 | None |
| validations | validations: 10 → 10 | None |
| conditional-formats | conditionalFormats: 21 → 21 | None |
| images | media: 7 → 5 | `xl/media/image.svg`, `xl/media/image2.svg` |
| shapes |  | None |
| chart-settings | charts: 1 → 1 | None |
| names | names: 2 → 2 | None |
| named-function-variants | names: 3 → 3 | None |
| data-table |  | None |
| one-variable-data-tables |  | None |
| pivot-slicer | pivots: 3 → 3; slicers: 2 → 1 | `xl/slicers/slicer.xml` |
| slicer-filter | pivots: 3 → 3; slicers: 2 → 1 | `xl/slicers/slicer.xml` |
| table-total-fresh-formula | tables: 1 → 1 | None |

## All 39 operation commands

Targets are concrete sheet names, IDs or range objects according to each command. Creation commands may accept `as` aliases, returned through `idMap`. These are workbook.apply commands; similarly named awareness commands can have different implementations.

### theme.colorScheme.set

Replace the workbook theme color scheme snapshot.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-theme.colorScheme.set-0):**

```json
{
  "op": "theme.colorScheme.set",
  "scheme": {
    "name": "Ocean",
    "themeColors": {
      "accent1": "#0F4C81",
      "accent2": "#00A6A6",
      "bg1": "#FFFFFF",
      "tx1": "#0F172A"
    }
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### sheet.add

Add a worksheet to the workbook.

Add a worksheet to the workbook.

**Required fields**
- `op`

**Optional fields**
- `as`
- `name`



Source: [sheet.md](files/sources/package/references/ops/sheet.md).

**Executed input (OP-sheet.add-0):**

```json
{
  "op": "sheet.add",
  "name": "Inventory"
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: false.
### sheet.set

Rename or reorder a worksheet.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-sheet.set-0):**

```json
{
  "op": "sheet.set",
  "target": "Sheet1",
  "props": {
    "name": "Dashboard",
    "index": 0
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### sheet.remove

Delete a worksheet by name.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-sheet.remove-0):**

```json
{
  "op": "sheet.remove",
  "target": "Scratch"
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### range.values.set

Set cell values for a range.

Set cell values for a range.

**Required fields**
- `op`
- `target`
- `values`

**Optional fields**
- (none)



Source: [range.md](files/sources/package/references/ops/range.md).

**Executed input (OP-range.values.set-0):**

```json
{
  "op": "range.values.set",
  "target": {
    "sheet": "Inventory",
    "range": "A1:C2"
  },
  "values": [
    [
      "SKU",
      "Color",
      "Stock"
    ],
    [
      "100-001",
      "Black",
      15
    ]
  ]
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### range.formulas.set

Set formulas for a range.

Set formulas for a range.

**Required fields**
- `formulas`
- `op`
- `target`

**Optional fields**
- (none)



Source: [range.md](files/sources/package/references/ops/range.md).

**Executed input (OP-range.formulas.set-0):**

```json
{
  "op": "range.formulas.set",
  "target": {
    "sheet": "Inventory",
    "range": "E2"
  },
  "formulas": [
    [
      "=AVERAGE(C2:C4)"
    ]
  ]
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### range.merge

Merge cells in a range.

Merge cells in a range.

**Required fields**
- `op`
- `target`

**Optional fields**
- `across`



Source: [range.md](files/sources/package/references/ops/range.md).

**Executed input (OP-range.merge-0):**

```json
{
  "op": "range.merge",
  "target": {
    "sheet": "Report",
    "range": "A1:D1"
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### range.unmerge

Unmerge any merged blocks intersecting the range.

Unmerge any merged blocks intersecting the range.

**Required fields**
- `op`
- `target`

**Optional fields**
- (none)



Source: [range.md](files/sources/package/references/ops/range.md).

**Executed input (OP-range.unmerge-0):**

```json
{
  "op": "range.unmerge",
  "target": {
    "sheet": "Report",
    "range": "A1:D4"
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### range.format.set

Apply formatting to a range.

Apply formatting to a range.

**Required fields**
- `op`
- `props`
- `target`

**Optional fields**
- `props.borders`
- `props.columnWidth`
- `props.fill`
- `props.font`
- `props.horizontalAlignment`
- `props.numberFormat`
- `props.rowHeight`
- `props.verticalAlignment`
- `props.wrapText`

**Enums**
- `props.borders.preset`: See [enums](files/sources/package/references/enums.md).
- `props.fill.color.value`: See [enums](files/sources/package/references/enums.md).
- `props.fill.gradientKind`: See [enums](files/sources/package/references/enums.md).
- `props.fill.pattern.type`: See [enums](files/sources/package/references/enums.md).
- `props.horizontalAlignment`: See [enums](files/sources/package/references/enums.md).
- `props.verticalAlignment`: See [enums](files/sources/package/references/enums.md).



Source: [range.md](files/sources/package/references/ops/range.md).

**Executed input (OP-range.format.set-0):**

```json
{
  "op": "range.format.set",
  "target": {
    "sheet": "Scorecard",
    "range": "A1:C1"
  },
  "props": {
    "fill": "accent1",
    "font": {
      "bold": true,
      "size": 14
    },
    "horizontalAlignment": "center",
    "numberFormat": "$#,##0.00",
    "wrapText": true,
    "borders": {
      "preset": "outside",
      "style": "solid",
      "color": "accent3"
    }
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### range.format.clear

Clear cell formatting for a range.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-range.format.clear-0):**

```json
{
  "op": "range.format.clear",
  "target": {
    "sheet": "Scorecard",
    "range": "A2:C8"
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### table.add

Create a table from a range.

Create a table from a range.

**Required fields**
- `op`
- `props`
- `props.range`

**Optional fields**
- `as`
- `props.hasHeaders`
- `props.name`



Source: [table.md](files/sources/package/references/ops/table.md).

**Executed input (OP-table.add-0):**

```json
{
  "op": "table.add",
  "props": {
    "range": {
      "sheet": "Inventory",
      "range": "A1:C3"
    },
    "hasHeaders": true,
    "name": "InventoryTable"
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### table.rows.add

Append or insert rows into a table.

Append or insert rows into a table.

**Required fields**
- `op`
- `props`
- `target`
- `props.values`

**Optional fields**
- `props.index`



Source: [table.md](files/sources/package/references/ops/table.md).

**Executed input (OP-table.rows.add-0):**

```json
{
  "op": "table.rows.add",
  "target": {
    "name": "InventoryTable",
    "sheet": "Inventory"
  },
  "props": {
    "index": null,
    "values": [
      [
        "100-003",
        "Denim",
        12
      ]
    ]
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### table.set

Update table metadata.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-table.set-0):**

```json
{
  "op": "table.set",
  "target": {
    "name": "InventoryTable",
    "sheet": "Inventory"
  },
  "props": {
    "name": "InventoryArchive"
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### table.remove

Delete a table from a worksheet.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-table.remove-0):**

```json
{
  "op": "table.remove",
  "target": {
    "name": "InventoryTable",
    "sheet": "Inventory"
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### conditionalformat.add

Add a conditional formatting rule to a range.

Add a conditional formatting rule to a range.

**Required fields**
- `op`
- `props`
- `target`
- `props.rule`

**Optional fields**
- (none)

**Enums**
- `props.rule.format.fill.color.value`: See [enums](files/sources/package/references/enums.md).
- `props.rule.format.fill.gradientKind`: See [enums](files/sources/package/references/enums.md).
- `props.rule.format.fill.pattern.type`: See [enums](files/sources/package/references/enums.md).
- `props.rule.operator`: See [enums](files/sources/package/references/enums.md).
- `props.rule.thresholds[].type`: See [enums](files/sources/package/references/enums.md).



Source: [conditionalformat.md](files/sources/package/references/ops/conditionalformat.md).

**Executed input (OP-conditionalformat.add-0):**

```json
{
  "op": "conditionalformat.add",
  "target": {
    "sheet": "CF",
    "range": "A1:A3"
  },
  "props": {
    "rule": {
      "type": "cellIs",
      "operator": "greaterThan",
      "formula": 3,
      "format": {
        "fill": "accent2",
        "font": {
          "bold": true,
          "color": "accent3"
        }
      }
    }
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.

**Executed input (OP-conditionalformat.add-1):**

```json
{
  "op": "conditionalformat.add",
  "target": {
    "sheet": "KPIs",
    "range": "F2:F20"
  },
  "props": {
    "rule": {
      "type": "top10",
      "rank": 10,
      "percent": true,
      "format": {
        "fill": "#DCFCE7",
        "font": {
          "color": "#166534"
        }
      }
    }
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.

**Executed input (OP-conditionalformat.add-2):**

```json
{
  "op": "conditionalformat.add",
  "target": {
    "sheet": "KPIs",
    "range": "G2:G20"
  },
  "props": {
    "rule": {
      "type": "iconSet",
      "iconSet": "5Rating",
      "showValue": false,
      "reverse": false,
      "percent": true
    }
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### conditionalformat.clear

Remove conditional formatting rules from a range.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-conditionalformat.clear-0):**

```json
{
  "op": "conditionalformat.clear",
  "target": {
    "sheet": "CF",
    "range": "A1:A3"
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### datavalidation.set

Set data validation rules for a range.

Set data validation rules for a range.

**Required fields**
- `op`
- `props`
- `target`
- `props.rule`

**Optional fields**
- `props.errorAlert`
- `props.ignoreBlanks`
- `props.inCellDropDown`
- `props.prompt`

**Enums**
- `props.errorAlert.style`: See [enums](files/sources/package/references/enums.md).
- `props.rule.operator`: See [enums](files/sources/package/references/enums.md).
- `props.rule.type`: See [enums](files/sources/package/references/enums.md).



Source: [datavalidation.md](files/sources/package/references/ops/datavalidation.md).

**Executed input (OP-datavalidation.set-0):**

```json
{
  "op": "datavalidation.set",
  "target": {
    "sheet": "Validation",
    "range": "B2:B10"
  },
  "props": {
    "rule": {
      "type": "list",
      "values": [
        "Dog",
        "Cat",
        "Bat"
      ]
    },
    "prompt": {
      "title": "Pick an animal",
      "show": true
    },
    "errorAlert": {
      "title": "Invalid choice",
      "show": true
    },
    "ignoreBlanks": true,
    "inCellDropDown": true
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### datavalidation.clear

Remove data validation rules from a range.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-datavalidation.clear-0):**

```json
{
  "op": "datavalidation.clear",
  "target": {
    "sheet": "Validation",
    "range": "B2:B10"
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### sparkline.add

Create a sparkline group backed by source data.

Create a sparkline group backed by source data.

**Required fields**
- `op`
- `props`
- `props.sourceData`
- `props.targetRange`
- `props.type`

**Optional fields**
- `props.axis`
- `props.axisColor`
- `props.dateAxisRange`
- `props.displayEmptyCellsAs`
- `props.displayHidden`
- `props.firstMarkerColor`
- `props.highMarkerColor`
- `props.lastMarkerColor`
- `props.lineWeight`
- `props.lowMarkerColor`
- `props.markers`
- `props.markersColor`
- `props.negativeColor`
- `props.seriesColor`

**Enums**
- `props.axis.maxMode`: See [enums](files/sources/package/references/enums.md).
- `props.axis.minMode`: See [enums](files/sources/package/references/enums.md).
- `props.displayEmptyCellsAs`: See [enums](files/sources/package/references/enums.md).
- `props.seriesColor.value`: See [enums](files/sources/package/references/enums.md).
- `props.type`: See [enums](files/sources/package/references/enums.md).



Source: [sparkline.md](files/sources/package/references/ops/sparkline.md).

**Executed input (OP-sparkline.add-0):**

```json
{
  "op": "sparkline.add",
  "props": {
    "uid": "{8FA89B8F-27A7-4E3D-A756-4CD08F8D7B2B}",
    "type": "line",
    "targetRange": {
      "sheet": "KPIs",
      "range": "H2:H3"
    },
    "sourceData": {
      "sheet": "KPIs",
      "range": "B2:G3"
    },
    "dateAxisRange": {
      "sheet": "KPIs",
      "range": "B1:G1"
    },
    "seriesColor": "accent1",
    "markers": {
      "show": true,
      "high": true,
      "low": true
    }
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### sparkline.set

Replace a sparkline group's recorded settings.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-sparkline.set-0):**

```json
{
  "op": "sparkline.set",
  "target": {
    "sheet": "KPIs",
    "uid": "{8FA89B8F-27A7-4E3D-A756-4CD08F8D7B2B}"
  },
  "props": {
    "type": "line",
    "targetRange": {
      "sheet": "KPIs",
      "range": "I2:I5"
    },
    "sourceData": {
      "sheet": "KPIs",
      "range": "C2:H5"
    },
    "displayHidden": true,
    "axis": {
      "showAxis": true,
      "rightToLeft": true
    }
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### sparkline.remove

Delete a sparkline group by uid.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-sparkline.remove-0):**

```json
{
  "op": "sparkline.remove",
  "target": {
    "sheet": "KPIs",
    "uid": "{8FA89B8F-27A7-4E3D-A756-4CD08F8D7B2B}"
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### chart.add

Add a chart anchored to worksheet cells.

Add a chart anchored to worksheet cells.

**Required fields**
- `op`
- `props`
- `sheet`
- `props.anchor`
- `props.chartType`

**Optional fields**
- `as`
- `props.categories`
- `props.dataLabels`
- `props.displayBlanksAs`
- `props.hasLegend`
- `props.legend`
- `props.series`
- `props.title`

**Enums**
- `props.chartType`: See [enums](files/sources/package/references/enums.md).
- `props.displayBlanksAs`: See [enums](files/sources/package/references/enums.md).
- `props.legend.position`: See [enums](files/sources/package/references/enums.md).
- `props.series[].marker.symbol`: See [enums](files/sources/package/references/enums.md).
- `props.series[].stroke.fill.color.value`: See [enums](files/sources/package/references/enums.md).
- `props.series[].stroke.fill.gradientKind`: See [enums](files/sources/package/references/enums.md).
- `props.series[].stroke.fill.pattern.type`: See [enums](files/sources/package/references/enums.md).
- `props.series[].stroke.style`: See [enums](files/sources/package/references/enums.md).



Source: [chart.md](files/sources/package/references/ops/chart.md).

**Executed input (OP-chart.add-0):**

```json
{
  "op": "chart.add",
  "sheet": "Charts",
  "props": {
    "chartType": "line",
    "anchor": {
      "from": {
        "row": 1,
        "col": 1,
        "rowOffsetPx": 4,
        "colOffsetPx": 8
      },
      "extent": {
        "widthPx": 520,
        "heightPx": 280
      }
    },
    "title": "Milky Way Star Birth Rate",
    "categories": [
      "2020",
      "2021",
      "2022",
      "2023"
    ],
    "series": [
      {
        "name": "Milky Way",
        "values": [
          1.8,
          1.9,
          2,
          2.2
        ]
      }
    ],
    "hasLegend": true,
    "legend": {
      "position": "bottom"
    },
    "dataLabels": {
      "showValue": true
    }
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### chart.set

Replace an existing chart's workbook-facing snapshot.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-chart.set-0):**

```json
{
  "op": "chart.set",
  "target": {
    "sheet": "Charts",
    "selector": {
      "chartType": "line",
      "title": "Updated chart"
    }
  },
  "props": {
    "chartType": "line",
    "anchor": {
      "from": {
        "row": 1,
        "col": 5
      },
      "extent": {
        "widthPx": 520,
        "heightPx": 280
      }
    },
    "title": "Updated chart",
    "categories": [
      "Q1",
      "Q2",
      "Q3"
    ],
    "series": [
      {
        "name": "Revenue",
        "values": [
          10,
          12,
          15
        ]
      }
    ]
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### chart.remove

Delete a worksheet chart by selector or id.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-chart.remove-0):**

```json
{
  "op": "chart.remove",
  "target": {
    "sheet": "Charts",
    "selector": {
      "chartType": "line",
      "title": "Updated chart"
    }
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### shape.add

Add a shape anchored to worksheet cells.

Add a shape anchored to worksheet cells.

**Required fields**
- `op`
- `props`
- `sheet`
- `props.geometry`

**Optional fields**
- `as`
- `props.anchor`
- `props.fill`
- `props.line`

**Enums**
- `props.fill.color.value`: See [enums](files/sources/package/references/enums.md).
- `props.fill.gradientKind`: See [enums](files/sources/package/references/enums.md).
- `props.fill.pattern.type`: See [enums](files/sources/package/references/enums.md).
- `props.geometry`: See [enums](files/sources/package/references/enums.md).
- `props.line.style`: See [enums](files/sources/package/references/enums.md).



Source: [shape.md](files/sources/package/references/ops/shape.md).

**Executed input (OP-shape.add-0):**

```json
{
  "op": "shape.add",
  "sheet": "Shapes",
  "props": {
    "geometry": "rect",
    "anchor": {
      "from": {
        "row": 2,
        "col": 3
      },
      "extent": {
        "widthPx": 260,
        "heightPx": 140
      }
    },
    "fill": "accent1",
    "line": {
      "style": "dashed",
      "fill": "accent4",
      "width": 1
    }
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### shape.set

Update a worksheet shape's recorded properties.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-shape.set-0):**

```json
{
  "op": "shape.set",
  "target": {
    "sheet": "Shapes",
    "selector": {
      "geometry": "rect"
    }
  },
  "props": {
    "fill": "accent2",
    "line": {
      "style": "solid",
      "fill": "accent3",
      "width": 2
    }
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### shape.remove

Delete a worksheet shape by selector or id.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-shape.remove-0):**

```json
{
  "op": "shape.remove",
  "target": {
    "sheet": "Shapes",
    "selector": {
      "geometry": "rect"
    }
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### image.add

Add an image anchored to worksheet cells.

Add an image anchored to worksheet cells.

**Required fields**
- `op`
- `props`
- `sheet`

**Optional fields**
- `as`
- `props.alt`
- `props.anchor`
- `props.contentType`
- `props.dataUrl`
- `props.path`
- `props.prompt`
- `props.uri`



Source: [image.md](files/sources/package/references/ops/image.md).

**Executed input (OP-image.add-0):**

```json
{
  "op": "image.add",
  "sheet": "Images",
  "props": {
    "dataUrl": "data:image/png;base64,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",
    "anchor": {
      "from": {
        "row": 1,
        "col": 2
      },
      "extent": {
        "widthPx": 160,
        "heightPx": 120
      }
    }
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.

**Adaptation:** Replaced placeholder data URL with local synthetic PNG.
### image.set

Update an existing worksheet image reference or anchor.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-image.set-0):**

```json
{
  "op": "image.set",
  "target": {
    "sheet": "Images",
    "imageId": "83af3002-5e7f-4ab6-88c6-f1e67c7be7b9",
    "anchor": {
      "from": {
        "row": 1,
        "col": 2
      },
      "extent": {
        "widthPx": 160,
        "heightPx": 120
      }
    }
  },
  "props": {
    "anchor": {
      "from": {
        "row": 3,
        "col": 4
      },
      "extent": {
        "widthPx": 180,
        "heightPx": 140
      }
    }
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.

**Adaptation:** Resolved actual image asset ID.
### image.remove

Delete a worksheet image by asset id and anchor.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-image.remove-0):**

```json
{
  "op": "image.remove",
  "target": {
    "sheet": "Images",
    "imageId": "ef932ac9-4921-450a-81b8-84982177e80e",
    "anchor": {
      "from": {
        "row": 1,
        "col": 2
      },
      "extent": {
        "widthPx": 160,
        "heightPx": 120
      }
    }
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.

**Adaptation:** Resolved actual image asset ID.
### names.range.add

Define a workbook or worksheet named range.

Define a workbook or worksheet named range.

**Required fields**
- `formula`
- `name`
- `op`

**Optional fields**
- `description`
- `sheet`



Source: [names.md](files/sources/package/references/ops/names.md).

**Executed input (OP-names.range.add-0):**

```json
{
  "op": "names.range.add",
  "name": "SalesRange",
  "formula": "Data!$A$1:$A$3",
  "description": "Input sales data"
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### names.function.add

Define a workbook or worksheet named function (LAMBDA).

Define a workbook or worksheet named function (LAMBDA).

**Required fields**
- `lambda`
- `name`
- `op`

**Optional fields**
- `description`
- `parameters`
- `returns`
- `sheet`



Source: [names.md](files/sources/package/references/ops/names.md).

**Executed input (OP-names.function.add-0):**

```json
{
  "op": "names.function.add",
  "name": "AddTax",
  "lambda": "LAMBDA(price, tax, price*(1+tax))",
  "description": "Adds tax to a price",
  "parameters": [
    {
      "name": "price",
      "description": "Base price"
    },
    {
      "name": "tax",
      "description": "Tax rate (e.g. 0.1)"
    }
  ],
  "returns": "Taxed price"
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### names.remove

Delete a workbook or worksheet scoped defined name.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-names.remove-0):**

```json
{
  "op": "names.remove",
  "name": "SalesRange"
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.

**Executed input (OP-names.remove-1):**

```json
{
  "op": "names.remove",
  "name": "LocalRange",
  "sheet": "Sheet1"
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### comments.self.set

Set the current comment author identity.

Set the current comment author identity.

**Required fields**
- `op`
- `person`

**Optional fields**
- (none)



Source: [comments.md](files/sources/package/references/ops/comments.md).

**Executed input (OP-comments.self.set-0):**

```json
{
  "op": "comments.self.set",
  "person": {
    "displayName": "Artifact Bot",
    "initials": "AB"
  }
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### thread.add

Create a threaded comment on a cell or range.

Create a threaded comment on a cell or range.

**Required fields**
- `body`
- `op`
- `target`

**Optional fields**
- `author`
- `createdAt`



Source: [thread.md](files/sources/package/references/ops/thread.md).

**Executed input (OP-thread.add-0):**

```json
{
  "op": "thread.add",
  "target": {
    "cell": {
      "sheet": "Sheet1",
      "address": "F5"
    }
  },
  "body": "Cell comment"
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.
### thread.reply

Reply to a comment thread.

Reply to a comment thread.

**Required fields**
- `body`
- `op`
- `target`

**Optional fields**
- `author`
- `createdAt`



Source: [thread.md](files/sources/package/references/ops/thread.md).

**Executed input (OP-thread.reply-0):**

```json
{
  "op": "thread.reply",
  "target": "{5E8825DC-66BB-4EC5-9C70-BE97A91D09E6}",
  "body": "Reply on cell"
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.

**Adaptation:** Resolved actual thread ID.
### thread.resolve

Resolve a comment thread.

Resolve a comment thread.

**Required fields**
- `op`
- `target`

**Optional fields**
- (none)



Source: [thread.md](files/sources/package/references/ops/thread.md).

**Executed input (OP-thread.resolve-0):**

```json
{
  "op": "thread.resolve",
  "target": "{65E93783-C04B-4AAC-B539-326C03F37E67}"
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.

**Adaptation:** Resolved actual thread ID.
### thread.reopen

Reopen a resolved comment thread.

Reopen a resolved comment thread.

**Required fields**
- `op`
- `target`

**Optional fields**
- (none)



Source: [thread.md](files/sources/package/references/ops/thread.md).

**Executed input (OP-thread.reopen-0):**

```json
{
  "op": "thread.reopen",
  "target": "{E7FC24CF-6DC1-4441-AEF2-9AE4C4E17FFA}"
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.

**Adaptation:** Resolved actual thread ID.
### thread.remove

Delete a threaded comment.

The help entry supplies examples but no separate required/optional-field schema in the captured operation references.

**Executed input (OP-thread.remove-0):**

```json
{
  "op": "thread.remove",
  "target": "th/{B2B08D4A-848D-4D8E-B6F7-CFD58DA573ED}"
}
```

**Output:** `{"idMap":{},"warnings":[]}`; serialized state changed: true.

**Adaptation:** Resolved actual thread ID.

## All 92 API help records and 126 examples

The help catalogue includes aliases and repeated topics, so 92 records is not a count of distinct methods. The harness provided workbook/sheet/range/chart variables and seeded prerequisite sheets. It removed TypeScript non-null assertions where necessary, converted structured examples to method calls and replaced placeholder IDs. The initial run is retained separately. Four examples execute but fail the export checkpoint; corrected facade usage is in the main guide.

### API 01 · charts

Chart objects anchored to cells.

- Use worksheet.charts.add(...) to create charts.
- Use worksheet.charts.deleteAll() (or .clear()) to remove charts.
- Use chart.delete() to remove a single chart.
- Charts are anchored with 0-based row/col positions and px offsets.
- Charts support title.text, setPosition, width/height, and axes.valueAxis/categoryAxis.
- Passing a multi-cell range to setPosition reserves that rectangle (from/to anchors).
### API 02 · worksheet.charts.add

Add a chart using a config object.

- Prefer passing a single config object to worksheet.charts.add(...) or assigning chart.xAxis/chart.yAxis/chart.legend/chart.dataLabels/chart.dataTable with compact config objects.
- Config assignment merges; omitted fields keep their existing values (useful for incremental edits).
- Skip noisy defaults in configs (e.g. text color is already theme text). Set only what changes the output.
- Use either manual anchors ({ from, to/extent }) or the compat range overload.
- Common chart types used in specs: 'bar', 'line', 'scatter'.
- Config props used in specs: title, titleTextStyle, categories, series, hasLegend, legend, barOptions, xAxis, yAxis, dataLabels, dataTable.
- Charts use 0-based row/col anchors for manual positioning.
- Charts support title.text, setPosition, width/height, and axes.valueAxis/categoryAxis.

<details>
<summary>Example 0: Create a chart with a manual anchor — executes and exports</summary>

```js
const sheet = workbook.worksheets.add('Charts');
const chart = sheet.charts.add('line', {
  title: 'Milky Way Star Birth Rate',
  titleTextStyle: { fontSize: 18, bold: true },
  categories: ['2020', '2021', '2022'],
  series: [
    { name: 'Milky Way', values: [1.8, 1.9, 2.0] },
  ],
  hasLegend: true,
  legend: { position: 'top' },
  barOptions: { direction: 'column', grouping: 'stacked', gapWidth: 120 },
  dataLabels: { showValue: false },
  from: { row: 1, col: 1 },
  extent: { widthPx: 520, heightPx: 280 },
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-000.json).

</details>


<details>
<summary>Example 1: Compat range-anchored chart API — executes and exports</summary>

```js
const sheet = workbook.worksheets.getItem('Charts');
const range = sheet.getRange('A2:B5');
const chart = sheet.charts.add('ColumnClustered', range, 'Auto');
chart.title.text = 'Stock Prices and Market Caps';
chart.setPosition(sheet.getRange('F2:M10'));
chart.width = 500;
chart.height = 300;
chart.axes.valueAxis.majorGridlines.format.line.color = '#CCCCCC';
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-001.json).

</details>

### API 03 · charts.add

Alias for worksheet.charts.add.

- Use worksheet.charts.add for the full signature and examples.
### API 04 · chart data table

Chart data table visibility and styling.

<details>
<summary>Example 2: Data table with optional legend keys — executes; export fails</summary>

```js
chart.dataTable = { visible: true, showLegendKey: true, textStyle: { fontSize: 10 } };
chart.dataTable = { visible: true, showLegendKey: false };
```

Result: TypeError: this[#C].toProto is not a function

[Full receipt](files/runs/api-example-002.json).

</details>

### API 05 · chart series

Series data, styling, and trendlines for worksheet charts.

- Use chart.series.add(name) then set categories/values.
- Trendlines are configured on a series via series.trendlines.add(...).

<details>
<summary>Example 3: Series + moving-average trendline — executes and exports</summary>

```js
const sheet = workbook.worksheets.add('Trendlines');
const chart = sheet.charts.add('line', { from: { row: 1, col: 1 }, extent: { widthPx: 420, heightPx: 240 } });
chart.categories = ['Q1', 'Q2', 'Q3', 'Q4'];
const series = chart.series.add('Revenue');
series.values = [10, 14, 13, 17];
series.categories = chart.categories;
series.trendlines.add('movingAverage', { movingAveragePeriod: 3 });
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-003.json).

</details>

### API 06 · chart title

Chart title text, placement, and text style.

<details>
<summary>Example 4: Compact title edits — executes and exports</summary>

```js
chart.title = '';
chart.titlePlacement = 'none';
chart.title = 'Sales Data';
chart.titlePlacement = 'aboveChart';
chart.titlePlacement = 'centeredOverlay';
chart.titleTextStyle.bold = true;
chart.titleTextStyle.fontSize = 18;
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-004.json).

</details>

### API 07 · chart fills

Chart and plot area fills (including pattern fills).

<details>
<summary>Example 5: Pattern fills (+ Office.js-style aliases) — executes and exports</summary>

```js
chart.chartFill = { type: 'solid', color: 'background1', pattern: { type: 'percent10', color: 'accent1' } };
chart.plotAreaFill = { type: 'solid', color: 'background1', pattern: { type: 'diagonalCross', color: 'accent2' } };

// Office.js-style aliases (map to chartFill/plotAreaFill + chartLine/plotAreaLine):
chart.chartArea.format.fill = { type: 'solid', color: 'background1' };
chart.plotArea.format.fill.setSolidColor('accent2');
chart.plotArea.format.line.visible = false;
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-005.json).

</details>

### API 08 · chart legend

Legend visibility, position, and styling.

<details>
<summary>Example 6: Legend on/off + position + style — executes and exports</summary>

```js
chart.hasLegend = false;
chart.hasLegend = true; chart.legend = { position: 'bottom', textStyle: { fontSize: 11 }, fill: { type: 'solid', color: 'background1', pattern: { type: 'zigZag', color: 'accent2' } }, stroke: { fill: 'accent2', style: 'solid', width: 1 } };
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-006.json).

</details>

### API 09 · chart barOptions

Bar/column chart layout options.

<details>
<summary>Example 7: Column chart spacing + overlap — executes and exports</summary>

```js
chart.barOptions.direction = 'column';
chart.barOptions.grouping = 'clustered';
chart.barOptions.gapWidth = 100;
chart.barOptions.overlap = -20;
chart.barOptions.direction = 'column';
chart.barOptions.grouping = 'stacked';
chart.barOptions.gapWidth = 120;
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-007.json).

</details>

### API 10 · chart axes

Axis config (type, position, titles, ticks, gridlines, formats, and styling).

- Use chart.xAxis and chart.yAxis to configure the primary axes.
- Axis config setters merge; apply small patches without resetting other axis fields.

<details>
<summary>Example 8: Axis type + presence — executes and exports</summary>

```js
chart.xAxis = { axisType: 'textAxis' };
chart.xAxis = { axisType: 'dateAxis' };
chart.xAxis = { deleted: true, title: 'Years' };
chart.yAxis = { deleted: true, title: 'Revenue' };
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-008.json).

</details>


<details>
<summary>Example 9: Axis titles + text styles (avoid redundant defaults) — executes and exports</summary>

```js
chart.xAxis = { title: { text: 'Category', textStyle: { fontSize: 14, bold: true } } };
chart.yAxis = { title: { text: 'Value', textStyle: { fontSize: 12, italic: true, fill: 'accent1' } } };
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-009.json).

</details>


<details>
<summary>Example 10: Axis ordering + crossing — executes and exports</summary>

```js
chart.xAxis = { position: 'bottom', orientation: 'maxMin' };
chart.yAxis = { position: 'left', crossBetween: 'between' };
chart.yAxis = { crosses: 'autoZero' };
chart.yAxis = { crossesAt: 2 };
chart.yAxis = { crosses: 'max' };
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-010.json).

</details>


<details>
<summary>Example 11: Tick labels — executes and exports</summary>

```js
chart.xAxis = { tickLabelPosition: 'nextTo' };
chart.xAxis = { tickLabelPosition: 'high' };
chart.xAxis = { tickLabelPosition: 'low' };
chart.xAxis = { tickLabelPosition: 'none', tickLabelInterval: 2, tickLabelDistanceFromAxis: 10 };
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-011.json).

</details>


<details>
<summary>Example 12: Tick marks + units — executes and exports</summary>

```js
chart.xAxis = { majorUnit: 5, minorUnit: 1 };
chart.xAxis = { majorTickMark: 'outside', minorTickMark: 'inside' };
chart.yAxis = { majorTickMark: 'cross', minorTickMark: 'none' };
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-012.json).

</details>


<details>
<summary>Example 13: Gridlines + axis lines — executes and exports</summary>

```js
chart.yAxis = { majorGridlines: { fill: 'background2', style: 'dashed', width: 1 }, minorGridlines: { fill: 'background2', style: 'dotted', width: 0.5 } };
chart.xAxis = { majorGridlines: { fill: 'background2', style: 'solid', width: 1 }, minorGridlines: { fill: 'background2', style: 'dotted', width: 0.5 } };
chart.xAxis = { line: { fill: 'background2', style: 'solid', width: 1 } };
chart.yAxis = { line: { fill: 'background2', style: 'solid', width: 1 } };

// Hide gridlines:
chart.yAxis = { majorGridlines: null, minorGridlines: null };
chart.xAxis = { majorGridlines: null, minorGridlines: null };
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-013.json).

</details>


<details>
<summary>Example 14: Number formats — executes and exports</summary>

```js
chart.xAxis = { numberFormatCode: '0', numberFormatSourceLinked: true };
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-014.json).

</details>


<details>
<summary>Example 15: Axis text style — executes and exports</summary>

```js
chart.xAxis = { textStyle: { fontSize: 10 } };
chart.yAxis = { textStyle: { fontSize: 1, fill: 'background1' } };
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-015.json).

</details>

### API 11 · categoryAxis

Alias for the chart category axis. In Granola, this maps to chart.xAxis.

- Prefer chart.xAxis = { ... } for config-style edits.
- Also available as chart.axes.categoryAxis.

<details>
<summary>Example 16: Show category axis tick labels — executes and exports</summary>

```js
chart.xAxis = { deleted: false, tickLabelPosition: 'nextToAxis' };
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-016.json).

</details>

### API 12 · valueAxis

Alias for the chart value axis. In Granola, this maps to chart.yAxis.

- Prefer chart.yAxis = { ... } for config-style edits.
- Also available as chart.axes.valueAxis.

<details>
<summary>Example 17: Format value axis gridlines — executes and exports</summary>

```js
chart.yAxis = { majorGridlines: { fill: 'background2', style: 'dashed', width: 1 } };
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-017.json).

</details>

### API 13 · chart.xAxis

Primary x-axis (category axis) config setter.

- Assign a config object: chart.xAxis = { ... }. Omitted fields keep current values.
- Compat name: categoryAxis.

<details>
<summary>Example 18: Make x-axis labels visible — executes and exports</summary>

```js
chart.xAxis = { visible: true, tickLabelPosition: 'nextToAxis' };
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-018.json).

</details>

### API 14 · chart.yAxis

Primary y-axis (value axis) config setter.

- Assign a config object: chart.yAxis = { ... }. Omitted fields keep current values.
- Compat name: valueAxis.

<details>
<summary>Example 19: Set y-axis title — executes and exports</summary>

```js
chart.yAxis = { title: { text: 'Revenue' } };
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-019.json).

</details>

### API 15 · chart dataLabels

Chart-level data labels (including callouts and pattern fills).

- Use chart.dataLabels for defaults applied to all series/points.
- For per-point customization, use series.dataLabelOverrides (see separate help entry).

<details>
<summary>Example 20: Positions + minimal style — executes; export fails</summary>

```js
chart.dataLabels = { showValue: false, position: undefined };
chart.dataLabels = { showValue: true, position: 'center', textStyle: { fontSize: 9 } };
chart.dataLabels = { showValue: true, position: 'inEnd' };
chart.dataLabels = { showValue: true, position: 'inBase' };
chart.dataLabels = { showValue: true, position: 'outEnd', textStyle: { bold: true, fontSize: 10 } };
```

Result: TypeError: this[#R].toProto is not a function

[Full receipt](files/runs/api-example-020.json).

</details>


<details>
<summary>Example 21: Data callout styling (fill + stroke) — executes; export fails</summary>

```js
chart.dataLabels = { showValue: true, position: 'dataCallout', fill: { type: 'solid', color: 'background1', pattern: { type: 'percent20', color: 'accent1' } }, stroke: { fill: 'accent1', style: 'solid', width: 1 } };
```

Result: TypeError: this[#R].toProto is not a function

[Full receipt](files/runs/api-example-021.json).

</details>

### API 16 · chart dataTable

Chart data table visibility and styling.

<details>
<summary>Example 22: Data table with optional legend keys — executes; export fails</summary>

```js
chart.dataTable = { visible: true, showLegendKey: true, textStyle: { fontSize: 10 }, fill: { type: 'solid', color: 'background1', pattern: { type: 'smallGrid', color: 'accent3' } } };
chart.dataTable = { visible: true, showLegendKey: false };
```

Result: TypeError: this[#C].toProto is not a function

[Full receipt](files/runs/api-example-022.json).

</details>

### API 17 · chart series

Series data, styling, number formats, error bars, trendlines, and label overrides.

<details>
<summary>Example 23: Series styling via worksheet.charts.add config — executes and exports</summary>

```js
const chart = worksheet.charts.add('bar', { categories: ['Q1','Q2'], series: [{ name: '2025', values: [18, 14], line: { fill: '#7c3aed', style: 'solid', width: 2 }, fill: { type: 'solid', color: '#ede9fe', pattern: { type: 'smallCheck', color: '#8b5cf6' } } }, { name: '2026', values: [21, 17], line: { fill: '#d97706', style: 'solid', width: 2 }, fill: { type: 'solid', color: '#fef3c7', pattern: { type: 'dottedDiamond', color: '#f59e0b' } } }] });
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-023.json).

</details>


<details>
<summary>Example 24: Scatter series with xValues (config) — executes and exports</summary>

```js
const chart = worksheet.charts.add('scatter', { series: [{ name: 'S1', xValues: [1, 2, 3, 4], values: [2, 4, 3, 5] }] });
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-024.json).

</details>


<details>
<summary>Example 25: Create a series and set categories/values (one-liners) — executes and exports</summary>

```js
chart.categories = ['A', 'B', 'C'];
const series = chart.series.add('S1'); series.categories = chart.categories; series.values = [1, 2, 3];
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-025.json).

</details>


<details>
<summary>Example 26: Series value number format code — executes and exports</summary>

```js
chart.series.items[0]!.valuesFormatCode = '$0.00';
```

Adaptations: Removed TypeScript non-null assertion after array access.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-026.json).

</details>


<details>
<summary>Example 27: Error bars — executes and exports</summary>

```js
const series = chart.series.items[0]!;
series.errorBars = { type: 'standardError', line: { fill: 'accent1', style: 'solid', width: 1 } };
series.errorBars = { type: 'percentage', value: 5 };
series.errorBars = { type: 'standardDeviation' };
```

Adaptations: Removed TypeScript non-null assertion after array access.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-027.json).

</details>


<details>
<summary>Example 28: Trendlines — executes and exports</summary>

```js
const series = chart.series.items[0]!;
series.trendlines.add('linear', { forecastForward: 2, forecastBackward: 1, intercept: 0, displayEquation: true, displayRSquared: true, stroke: { fill: 'accent1', style: 'dashed', width: 1 } });
series.trendlines.add('polynomial', { polynomialOrder: 3 });
series.trendlines.add('movingAverage', { movingAveragePeriod: 5 });
series.trendlines.add('exponential');
series.trendlines.clear();
```

Adaptations: Removed TypeScript non-null assertion after array access.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-028.json).

</details>


<details>
<summary>Example 29: Trendline styling (label fill + line alias) — executes and exports</summary>

```js
const trend = series.trendlines.add('linear', { line: { fill: 'accent5', style: 'dashed', width: 1 }, label: { fill: { type: 'solid', color: 'background1', pattern: { type: 'smallCheck', color: 'accent6' } } } });
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-029.json).

</details>


<details>
<summary>Example 30: Per-point data label overrides — executes and exports</summary>

```js
const series = chart.series.items[1]!;
const label0 = series.dataLabelOverrides.add(0);
label0.showValue = true;
label0.position = 'inEnd';
label0.textStyle.bold = true;
label0.textStyle.fontSize = 10;
```

Adaptations: Removed TypeScript non-null assertion after array access.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-030.json).

</details>


<details>
<summary>Example 31: Series styling after creation (one-liners) — executes and exports</summary>

```js
chart.series.items[0]!.line = { fill: '#3b82f6', style: 'solid', width: 2 };
chart.series.items[0]!.fill = { type: 'solid', color: '#60a5fa' };
```

Adaptations: Removed TypeScript non-null assertion after array access.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-031.json).

</details>


<details>
<summary>Example 32: Outline-only bars (no fill + visible stroke) — executes and exports</summary>

```js
const series = chart.series.items[0]!;
series.fill = { type: 'none' };
series.stroke = { color: 'accent5', style: 'solid', weight: 1.5 };

// Equivalent format-style surface (useful for agent code):
series.format.fill.clear();
series.format.line = { color: 'accent5', style: 'solid', weight: 1.5 };
```

Adaptations: Removed TypeScript non-null assertion after array access.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-032.json).

</details>

### API 18 · conditional-formatting

Conditional formatting rules for ranges.

- Supported authored rule families include: cellIs, expression (Custom), containsText/notContainsText, beginsWith/endsWith, containsBlanks/notContainsBlanks, containsErrors/notContainsErrors, duplicateValues/uniqueValues, timePeriod, top10, aboveAverage, colorScale, dataBar, and iconSet.
- colorScale accepts { colors, thresholds } or criteria objects/arrays.
- iconSet accepts built-in Excel icon names such as 3Arrows, 3TrafficLights1, 4Rating, and 5Boxes, and it auto-generates 3/4/5 thresholds when omitted.
- Use range.conditionalFormats.add(type, config) to add rules.
- Use range.conditionalFormats.clear() (or deleteAll()) to remove rules.
- Use range.conditionalFormats.add('Custom', { formula, format }) for expression rules.
- Convenience aliases: range.conditionalFormats.addExpression(...) and addCustom(...).
- Compat alias: use type='CellValue' for cellIs rules.
- Use range.conditionalFormats.addCellIs({ operator, formula, format }) for cellIs rules.
### API 19 · range.conditionalFormats

Conditional formatting rules scoped to a range.

- Compat alias: use type='CellValue' for cellIs rules.
- Use Custom as an alias for expression rules.
- Supported rule types include: cellIs, colorScale, dataBar, iconSet, expression (Custom), text rules, blank/error rules, duplicate/unique, timePeriod, top10, and aboveAverage.
- colorScale accepts { colors, thresholds } or criteria objects/arrays.
- iconSet preserves showValue/reverse/custom/percent/thresholds in inspect output.
- Use range.conditionalFormats.add(type, config) to add rules.
- Use range.conditionalFormats.clear() (or deleteAll()) to remove rules.
- Use range.conditionalFormats.addCellIs({ operator, formula, format }) for cellIs rules.
- Use range.conditionalFormats.addExpression(...) and addCustom(...) as expression-rule convenience aliases.

<details>
<summary>Example 33: Add a color scale rule — executes and exports</summary>

```js
const range = sheet.getRange('B2:B6');
range.conditionalFormats.add('colorScale', {
  colors: ['#FCA5A5', '#DC2626'],
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-033.json).

</details>


<details>
<summary>Example 34: Add a 3-stop color scale (min/mid/max) — executes and exports</summary>

```js
const range = sheet.getRange('B2:B6');
range.conditionalFormats.add('colorScale', {
  criteria: [
    { type: 'lowestValue', color: '#DC2626' },
    { type: 'percentile', value: 50, color: '#FBBF24' },
    { type: 'highestValue', color: '#16A34A' },
  ],
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-034.json).

</details>


<details>
<summary>Example 35: Add a custom formula rule — executes and exports</summary>

```js
const range = sheet.getRange('B2:B6');
range.conditionalFormats.add('Custom', {
  formula: '=B2>10',
  format: { fill: '#FDE68A', font: { bold: true } },
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-035.json).

</details>


<details>
<summary>Example 36: Add a cell value rule — executes and exports</summary>

```js
const range = sheet.getRange('B2:B6');
range.conditionalFormats.addCellIs({
  operator: 'lessThan',
  formula: 0,
  format: { fill: '#F8CBAD' },
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-036.json).

</details>


<details>
<summary>Example 37: Add data bars (color bars) for magnitude — executes and exports</summary>

```js
const range = sheet.getRange('C2:C10');
range.conditionalFormats.add('dataBar', {
  color: 'accent5',
  gradient: true,
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-037.json).

</details>


<details>
<summary>Example 38: Add an icon set for quick up/flat/down signals — executes and exports</summary>

```js
const range = sheet.getRange('D2:D10');
range.conditionalFormats.add('iconSet', {
  iconSet: '3Arrows',
  showValue: false,
  reverse: false,
  percent: true,
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-038.json).

</details>


<details>
<summary>Example 39: Highlight text cells that contain a substring — executes and exports</summary>

```js
const range = sheet.getRange('E2:E8');
range.conditionalFormats.add('containsText', {
  text: 'late',
  format: { fill: '#FDE68A', font: { color: '#92400E', bold: true } },
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-039.json).

</details>


<details>
<summary>Example 40: Highlight dates in a relative time window — executes and exports</summary>

```js
const range = sheet.getRange('F2:F12');
range.conditionalFormats.add('timePeriod', {
  timePeriod: 'last7Days',
  format: { fill: '#DBEAFE', font: { color: '#1D4ED8' } },
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-040.json).

</details>


<details>
<summary>Example 41: Highlight top-ranked values by count or percent — executes and exports</summary>

```js
const range = sheet.getRange('G2:G20');
range.conditionalFormats.add('top10', {
  rank: 10,
  percent: true,
  format: { fill: '#DCFCE7', font: { color: '#166534' } },
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-041.json).

</details>


<details>
<summary>Example 42: Highlight values below or equal to the average — executes and exports</summary>

```js
const range = sheet.getRange('H2:H20');
range.conditionalFormats.add('aboveAverage', {
  aboveAverage: false,
  equalAverage: true,
  format: { fill: '#FECACA', font: { color: '#991B1B' } },
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-042.json).

</details>

### API 20 · range.conditionalFormats.addCustom

Add a custom (expression) conditional format to a range.

- addCustom(formula, format) takes the differential format as the second argument (fill/font/border/numberFormat).
- If you already have a { formula, format } object, use range.conditionalFormats.add('Custom', { ... }).

<details>
<summary>Example 43: Highlight cells using a custom formula — executes and exports</summary>

```js
const range = sheet.getRange('F2:F50');
range.conditionalFormats.addCustom('=F2<E2', {
  fill: '#FECACA',
  font: { color: '#991B1B', bold: true },
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-043.json).

</details>


<details>
<summary>Example 44: Equivalent add('Custom', ...) form — executes and exports</summary>

```js
const range = sheet.getRange('F2:F50');
range.conditionalFormats.add('Custom', {
  formula: '=F2<E2',
  format: { fill: '#FECACA', font: { color: '#991B1B' } },
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-044.json).

</details>

### API 21 · range.conditionalFormats.addExpression

Alias for addCustom; adds a custom (expression) conditional rule.

- Use range.conditionalFormats.addCustom(...) for the same behavior.
### API 22 · workbook

Workbook lifecycle and entry points.

- Create a workbook with Workbook.create(), add worksheets with workbook.worksheets.add(name), or load serialized data with Workbook.load().
- Access sheets via workbook.worksheets.
- Call workbook.recalculate() before reading formula results.
- Export with workbook.render(...) or workbook.toProto().
### API 23 · workbook.export

Export an editable XLSX, preview image, or structured layout data.

- Use workbook.export({ format: 'xlsx' }) for native Office output.
- Use workbook.render(...) for a PNG or JPEG preview.

<details>
<summary>Example 45: Export an editable workbook — executes and exports</summary>

```js
const xlsx = await workbook.export({ format: 'xlsx' });
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-045.json).

</details>

### API 24 · workbook.render

Render worksheet previews as PNG/JPEG.

- Use workbook.render({ sheetName, range, format, scale, headers }).
- range is A1 notation; center requires width and height.
- Provide only one of sheet, sheetName, or sheetIndex.
- Use range or center (with width and height).

<details>
<summary>Example 46: Render a range preview — executes and exports</summary>

```js
const png = await workbook.render({
  sheetName: 'Summary',
  range: 'A1:D12',
  format: 'png',
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-046.json).

</details>

### API 25 · workbook.help

Return NDJSON help about workbook APIs via query and regex search

- Tip: query matches entry names, not example bodies. 
- To find an API by keyword, use query='*' to grep across the help index and search=<regex>
- Search uses regex matching; for multiple terms use pipes (example: freeze|pane|freeze_panes).
- Use `include` param to control payload size.
- Formula metadata is exposed under fx.<NAME> entries (ex: fx.LET).
- Enums are listed under enum.<Name> (ex: enum.ShapeGeometry).

<details>
<summary>Example 47: Get help for a specific enums — executes and exports</summary>

```js
workbook.help(
  'enum.ShapeGeometry',
  { include: ['index','notes'] },
)
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-047.json).

</details>


<details>
<summary>Example 48: Get help for multiple enums — executes and exports</summary>

```js
workbook.help(
  'enum.*',
  { search: 'ShapeGeometry|LineStyle', include: ['index'] },
)
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-048.json).

</details>


<details>
<summary>Example 49: Get help on specific features by regex search — executes and exports</summary>

```js
workbook.help({query: '*', search: 'trendline|trendlines', include: ['index','examples']}).
```

Adaptations: Repaired malformed trailing dot and used the documented query/options signature.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-049.json).

</details>


<details>
<summary>Example 50: Range addressing and value/formula basics — executes and exports</summary>

```json
{
  "query": "ranges",
  "include": [
    "index",
    "examples"
  ],
  "search": null,
  "maxChars": 1200,
  "summary": "Range addressing and value/formula basics"
}
```

Adaptations: Mapped structured example fields to the corresponding public method; replaced placeholder anchors with actual fixture IDs.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-050.json).

Returned/logged output:

```text
{"ndjson":"{\"kind\":\"api\",\"name\":\"ranges\",\"summary\":\"Range addressing and cell reads/writes.\",\"tags\":[\"ranges\"]}","truncated":false,"metadata":{"revision":"o0lmjp","query":"ranges","include":{"requested":["index","examples"],"tokens":["index","examples"]},"notices":[]}}
```

</details>


<details>
<summary>Example 51: Formatting ranges (fills, fonts, borders, number formats) — executes and exports</summary>

```json
{
  "query": "styles",
  "include": [
    "index",
    "examples"
  ],
  "search": null,
  "maxChars": 1200,
  "summary": "Formatting ranges (fills, fonts, borders, number formats)"
}
```

Adaptations: Mapped structured example fields to the corresponding public method; replaced placeholder anchors with actual fixture IDs.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-051.json).

Returned/logged output:

```text
{"ndjson":"{\"kind\":\"api\",\"name\":\"styles\",\"summary\":\"Formatting ranges (fills, fonts, borders, number formats).\",\"tags\":[\"styles\",\"range.format\",\"formatting\",\"number format\",\"number_format\",\"autofit\",\"fill\",\"font\",\"borders\"]}\n{\"kind\":\"notice\",\"message\":\"Truncated: omitted 1 lines. Increase maxChars or narrow query.\"}","truncated":true,"metadata":{"revision":"8pwax9","query":"styles","include":{"requested":["index","examples"],"tokens":["index","examples"]},"notices":[]}}
```

</details>


<details>
<summary>Example 52: Grep for feature namespaces — executes and exports</summary>

```json
{
  "query": "*",
  "include": [
    "index",
    "examples",
    "notes"
  ],
  "search": "tables|sparklines|charts|images|render|validation|conditional|freeze|data_validation",
  "maxChars": 2000,
  "summary": "Grep for feature namespaces"
}
```

Adaptations: Mapped structured example fields to the corresponding public method; replaced placeholder anchors with actual fixture IDs.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-052.json).

Returned/logged output:

```text
{"ndjson":"{\"kind\":\"api\",\"name\":\"charts\",\"summary\":\"Chart objects anchored to cells.\",\"tags\":[\"charts\"],\"notes\":[\"Use worksheet.charts.add(...) to create charts.\",\"Use worksheet.charts.deleteAll() (or .clear()) to remove charts.\",\"Use chart.delete() to remove a single chart.\",\"Charts are anchored with 0-based row/col positions and px offsets.\",\"Charts support title.text, setPosition, width/height, and axes.valueAxis/categoryAxis.\",\"Passing a multi-cell range to setPosition reserves that rectangle (from/to anchors).\"]}\n{\"kind\":\"notice\",\"message\":\"Truncated: omitted 40 lines. Increase maxChars or narrow query.\"}","truncated":true,"metadata":{"revision":"6jpp5f","query":"*","include":{"requested":["index","examples","notes"],"tokens":["index","examples","notes"]},"search":"tables|sparklines|charts|images|render|validation|conditional|freeze|data_validation","notices":[]}}
```

</details>


<details>
<summary>Example 53: Browse formula metadata entries — executes and exports</summary>

```json
{
  "query": "fx.*",
  "include": [
    "index",
    "examples",
    "notes"
  ],
  "search": "LET|LAMBDA|XLOOKUP|MAP|BYROW",
  "maxChars": 2000,
  "summary": "Browse formula metadata entries"
}
```

Adaptations: Mapped structured example fields to the corresponding public method; replaced placeholder anchors with actual fixture IDs.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-053.json).

Returned/logged output:

```text
{"ndjson":"{\"kind\":\"formula\",\"name\":\"fx.EXPONDIST\",\"summary\":\"This function is available for compatibility with Excel 2007 and earlier. Returns the exponential distribution\",\"category\":\"compatibility\",\"examples\":[\"=EXPONDIST(A1, A1, A1)\"],\"notes\":[\"x (value|range, required): Is the value of the function, a nonnegative number\",\"lambda (value|range, required): Is the parameter value, a positive number\",\"cumulative (value|range, required): Is a logical value for the function to return: the cumulative distribution function = TRUE; the probability density function = FALSE\"]}\n{\"kind\":\"formula\",\"name\":\"fx.ISOMITTED\",\"summary\":\"Checks whether the value is omitted, and returns TRUE or FALSE\",\"category\":\"information\",\"examples\":[\"=ISOMITTED(A1)\"],\"notes\":[\"argument (value|range, required): Is the value you want to test, such as a LAMBDA parameter\"]}\n{\"kind\":\"formula\",\"name\":\"fx.LET\",\"summary\":\"Assigns calculation results to names. Useful for storing intermediate calculations and values by defining names inside a formula. These names only apply within the scope of the LET function.\",\"category\":\"logical\",\"examples\":[\"=LET(A1, A1, A1, C1:C5)\"],\"notes\":[\"name1 (value|range, required): The name, or a calculation which can make use of all names within the LET. Names must start with a letter, cannot be the output of a formula, or conflict with range syntax.\",\"nameValue1 (value|range, required): The value associated with the name.\",\"calculationOrName2 (value|range, required): Value, reference, or range.\",\"nameValue2 (value|range, optional): Value argument for the calculation.\",\"rest (value|range, required): Value, reference, or range.\"]}\n{\"kind\":\"notice\",\"message\":\"Truncated: omitted 13 lines. Increase maxChars or narrow query.\"}","truncated":true,"metadata":{"revision":"jvi7ad","query":"fx.*","include":{"requested":["index","examples","notes"],"tokens":["index","examples","notes"]},"search":"LET|LAMBDA|XLOOKUP|MAP|BYROW","notices":[]}}
```

</details>

### API 26 · workbook.resolve

Resolve anchored ids from inspect output.

- Worksheet anchors: ws/<sheetRef> where sheetRef can be a worksheet name, worksheet id, or sheetId.
- Thread anchors: th/<threadId>.
- Chart anchors from inspect use ch/<chartAnchor>. Legacy ch/<chartId> and ch/<sheetRef>.<chartId> forms are also accepted.

<details>
<summary>Example 54: workbook.resolve — executes and exports</summary>

```json
{
  "id": "ws/<sheetId>"
}
```

Adaptations: Mapped structured example fields to the corresponding public method; replaced placeholder anchors with actual fixture IDs.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-054.json).

Returned/logged output:

```text
vykaqn
```

</details>


<details>
<summary>Example 55: workbook.resolve — executes and exports</summary>

```json
{
  "id": "th/<threadId>"
}
```

Adaptations: Mapped structured example fields to the corresponding public method; replaced placeholder anchors with actual fixture IDs.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-055.json).

Returned/logged output:

```text
{DC989978-6189-4405-AEFF-C26ACDD46DB8}
```

</details>


<details>
<summary>Example 56: workbook.resolve — executes and exports</summary>

```json
{
  "id": "ch/<chartAnchor>"
}
```

Adaptations: Mapped structured example fields to the corresponding public method; replaced placeholder anchors with actual fixture IDs.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-056.json).

Returned/logged output:

```text
q0xdu5
```

</details>

### API 27 · workbook.recalculate

Recalculate formulas and computed values.

<details>
<summary>Example 57: Recalculate after writing formulas — executes and exports</summary>

```js
const sheet = workbook.worksheets.getItem('Sheet1');
sheet.getRange('A1:A2').values = [[1], [2]];
sheet.getRange('B1').formulas = [['=SUM(A:A)']];
const total = sheet.getRange('B1').values[0]?.[0];
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-057.json).

</details>

### API 28 · workbook.worksheets.add

Add one or more worksheets by name.

- If the name already exists, the existing worksheet is returned.
- Use workbook.worksheets.getItem to access existing sheets by name.

<details>
<summary>Example 58: Create a worksheet and write values — executes and exports</summary>

```js
const sheet = workbook.worksheets.add('DogHealth');
sheet.getRange('A1').values = [['Date']];
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-058.json).

</details>


<details>
<summary>Example 59: Create multiple worksheets — executes and exports</summary>

```js
const [dataSheet, summarySheet] = workbook.worksheets.add(['Data', 'Summary']);
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-059.json).

</details>

### API 29 · workbook.worksheets.getItem

Get a worksheet by name.

<details>
<summary>Example 60: Fetch a sheet and write a cell — executes and exports</summary>

```js
const sheet = workbook.worksheets.getItem('Sheet1');
sheet.getRange('A1').values = [['Hello']];
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-060.json).

</details>

### API 30 · workbook.worksheets.getActiveWorksheet

Return the active worksheet.

<details>
<summary>Example 61: Write to the active sheet — executes and exports</summary>

```js
const sheet = workbook.worksheets.getActiveWorksheet();
sheet.getRange('A1').values = [['Active sheet']];
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-061.json).

</details>

### API 31 · workbook.setColorScheme

Update the workbook theme color palette (accent1..6, bg/tx, dk/lt, hyperlinks).

- Theme slots: accent1..accent6, bg1/bg2, tx1/tx2, dk1/lt1/dk2/lt2, hlink, folHlink.
- Use theme tokens in styling APIs (fills, fonts, borders, charts) so global re-theming works.
- Hex colors are fine for special needs, but tokens are the best default for dashboards and consistent styling.

<details>
<summary>Example 62: Swap the workbook theme palette — executes and exports</summary>

```js
// Prefer updating the theme palette over repainting every range.
// Theme tokens like 'accent1' will automatically pick up the new palette.
workbook.setColorScheme({
  name: 'Brand',
  themeColors: {
    accent1: '#1F6FEB',
    accent2: '#F97316',
    accent3: '#16A34A',
    dk1: '#0B0F19',
    lt1: '#FFFFFF',
    lt2: '#E5E7EB',
    hlink: '#2563EB',
    folHlink: '#7C3AED',
  },
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-062.json).

</details>


<details>
<summary>Example 63: Style with theme tokens (accent slots) — executes and exports</summary>

```js
// Then use theme tokens for styling:
const sheet = workbook.worksheets.getItem('Sheet1');
const header = sheet.getRange('A1:D1');
header.format.fill = 'accent1';
header.format.font = { color: 'lt1', bold: true };
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-063.json).

</details>

### API 32 · workbook.utils

Small helpers for building formulas and addresses.

<details>
<summary>Example 64: Convert between 1-based columns and letters — executes and exports</summary>

```js
workbook.utils.columnToLetter(1); // "A"
workbook.utils.columnToLetter(27); // "AA"
workbook.utils.letterToColumn('AA'); // 27
workbook.utils.toA1String(1, 2); // 'B1'
workbook.utils.toA1String(1, 2, 3, 4); // 'B1:D3'
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-064.json).

</details>

### API 33 · workbook.awareness

Presence selections plus comment and note operations.

- Ops include comments.self.set, thread.add, thread.reply, note.add (comment.add alias).
- Comment/note ops require a comment author; call comments.self.set once or pass an author.
### API 34 · drawings

Charts, shapes, and images anchored to cells.

- Use worksheet.charts, worksheet.shapes, and worksheet.images.
- Anchors use 0-based row/col indices with px offsets.
- Clearing cells does not remove drawings; delete drawings explicitly.
### API 35 · worksheet.deleteAllDrawings

Remove all charts, shapes, and images from a worksheet.

<details>
<summary>Example 65: Clear drawings before rebuilding a dashboard — executes and exports</summary>

```js
sheet.deleteAllDrawings();
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-065.json).

</details>

### API 36 · worksheet.autoLayoutDrawings

Lay out charts/shapes/images deterministically in a frame.

- Layout inputs are pixels; items keep their existing sizes.
- Pass charts, shapes, or images (or any mix).

<details>
<summary>Example 66: Stack charts vertically in a dashboard frame — executes and exports</summary>

```js
const items = sheet.charts.items;
sheet.autoLayoutDrawings(items, {
  direction: 'vertical',
  frame: { startCell: 'A8', width: 1200, height: 800 },
  gap: 24,
  padding: 32,
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-066.json).

</details>

### API 37 · shapes

Shapes anchored to cells.

- Use worksheet.shapes.add(...) to create shapes.
- Shapes use 0-based row/col anchors and px offsets.
### API 38 · worksheet.shapes.add

Add a shape anchored to worksheet cells.

- geometry values are listed under enum.ShapeGeometry.
- geometry names are case-insensitive and accept common aliases like rectangle and rightTriangle.
- anchor accepts { from, to/extent } or a range like 'A1:D6'.

<details>
<summary>Example 67: Create a rectangle with an explicit extent anchor — executes and exports</summary>

```js
const sheet = workbook.worksheets.getItem('Shapes');
sheet.shapes.add({
  geometry: 'rect',
  anchor: {
    from: { row: 2, col: 3 },
    extent: { widthPx: 260, heightPx: 140 },
  },
  fill: 'accent1',
  line: { style: 'dashed', fill: 'accent4', width: 1 },
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-067.json).

</details>


<details>
<summary>Example 68: Compat add with a range anchor — executes and exports</summary>

```js
const sheet = workbook.worksheets.getItem('Shapes');
const star = sheet.shapes.add('Star5', sheet.getRange('G2:L10'));
star.name = 'Star';
star.fill.color = '#FFD966';
star.line.color = '#C09000';
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-068.json).

</details>

### API 39 · shapes.add

Alias for worksheet.shapes.add.

- Use worksheet.shapes.add for the full signature and examples.
### API 40 · formulas

Formula authoring and calculation.

- Write formulas via range.formulas (2D string array).
- range.values accepts strings starting with '=' as formulas; use "'=..." for literal text.
- Call workbook.recalculate() before reading computed values.
- If a newer Excel function appears unsupported in your runtime, replace it with compatible patterns (example: AVERAGEIFS -> SUMIFS/COUNTIFS).

<details>
<summary>Example 69: formulas — executes and exports</summary>

```js
const sheet = workbook.worksheets.add("Sheet1");
const range = sheet.getRange("A1:A2");
range.formulas = [["=1+1"], ["=SUM(2,3)"]];
workbook.recalculate();
const values = range.values;
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-069.json).

</details>

### API 41 · images

Worksheet image objects anchored to cells.

- Use worksheet.images.add(...) to place an image.
- Sources may use path, URI, dataUrl, inline svg, blob, or prompt.
- Anchors use 0-based row/col indices with px offsets.

<details>
<summary>Example 70: Place inline SVG without base64 encoding — executes and exports</summary>

```js
worksheet.images.add({
  svg: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" fill="#2563eb"/></svg>',
  anchor: { from: { row: 1, col: 1 }, extent: { widthPx: 48, heightPx: 48 } },
});
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-070.json).

</details>

### API 42 · workbook.inspect

Return NDJSON snapshot records for the workbook.

- Use search and maxChars to keep output focused.
- include/exclude select record properties (not cell addresses). Use search for A1-like addresses.
- Use searchTerm to return match records with sheet + address.
- range + sheetId scope table/formula/style output to a specific range.
- Use kind=conditionalFormatting (alias cf) to inspect rule families and params like text, timePeriod, top10, aboveAverage, and iconSet fields such as thresholds/showValue/reverse/percent.
- tableMaxRows/tableMaxCols/tableMaxCellChars control table preview sizes when values would be too large.
- Use kind=definedName to list workbook or sheet-scoped defined names.

<details>
<summary>Example 71: Inspect sheets/tables matching Revenue — executes and exports</summary>

```json
{
  "target": null,
  "kind": "sheet,table",
  "include": null,
  "exclude": null,
  "search": "Revenue",
  "maxChars": 1200,
  "searchTerm": null,
  "sheetId": null,
  "range": null,
  "offset": null,
  "options": null,
  "summary": "Inspect sheets/tables matching Revenue"
}
```

Adaptations: Mapped structured example fields to the corresponding public method; replaced placeholder anchors with actual fixture IDs.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-071.json).

Returned/logged output:

```text
{"recordCount":2,"truncated":false,"metadata":{"revision":"w68lpg","target":{"id":"wb/j9pl4e","beforeLines":0,"afterLines":0},"kind":{"requested":"sheet,table","tokens":["sheet","table"]},"include":{"tokens":[]},"exclude":{"tokens":[]},"search":"Revenue","notices":[]}}
```

</details>


<details>
<summary>Example 72: List data regions with previews for a sheet — executes and exports</summary>

```json
{
  "target": null,
  "kind": "region",
  "include": null,
  "exclude": null,
  "search": null,
  "maxChars": 2000,
  "tableMaxRows": 5,
  "tableMaxCols": 6,
  "tableMaxCellChars": 80,
  "searchTerm": null,
  "sheetId": "ws/<sheetId>",
  "range": null,
  "offset": null,
  "options": null,
  "summary": "List data regions with previews for a sheet"
}
```

Adaptations: Mapped structured example fields to the corresponding public method; replaced placeholder anchors with actual fixture IDs.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-072.json).

Returned/logged output:

```text
{"recordCount":6,"truncated":false,"metadata":{"revision":"c4bv3x","target":{"id":"wb/dz3h02","beforeLines":0,"afterLines":0},"kind":{"requested":"region","tokens":["region"]},"include":{"tokens":[]},"exclude":{"tokens":[]},"notices":["Worksheet <sheetId> not found."]}}
```

</details>


<details>
<summary>Example 73: Preview table values for a range — executes and exports</summary>

```json
{
  "target": null,
  "kind": "table",
  "include": "values",
  "exclude": null,
  "search": null,
  "maxChars": 2000,
  "tableMaxRows": 5,
  "tableMaxCols": 6,
  "tableMaxCellChars": 80,
  "searchTerm": null,
  "sheetId": "ws/<sheetId>",
  "range": "A1:D20",
  "offset": null,
  "options": null,
  "summary": "Preview table values for a range"
}
```

Adaptations: Mapped structured example fields to the corresponding public method; replaced placeholder anchors with actual fixture IDs.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-073.json).

Returned/logged output:

```text
{"recordCount":4,"truncated":true,"metadata":{"revision":"j7o7b2","target":{"id":"wb/eq7ekd","beforeLines":0,"afterLines":0},"kind":{"requested":"table","tokens":["table"]},"include":{"requested":"values","tokens":["values"]},"exclude":{"tokens":[]},"notices":["Worksheet <sheetId> not found."]}}
```

</details>


<details>
<summary>Example 74: List conditional formatting rules for a range — executes and exports</summary>

```json
{
  "target": null,
  "kind": "conditionalFormatting",
  "include": null,
  "exclude": null,
  "search": "top10",
  "maxChars": 2000,
  "tableMaxRows": 5,
  "tableMaxCols": 6,
  "tableMaxCellChars": 80,
  "searchTerm": null,
  "sheetId": "ws/<sheetId>",
  "range": "B2:B20",
  "offset": null,
  "options": null,
  "summary": "List conditional formatting rules for a range"
}
```

Adaptations: Mapped structured example fields to the corresponding public method; replaced placeholder anchors with actual fixture IDs.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-074.json).

Returned/logged output:

```text
{"recordCount":1,"truncated":false,"metadata":{"revision":"k10wvr","target":{"id":"wb/nhdvxh","beforeLines":0,"afterLines":0},"kind":{"requested":"conditionalFormatting","tokens":["conditionalFormatting"]},"include":{"tokens":[]},"exclude":{"tokens":[]},"search":"top10","notices":["Worksheet <sheetId> not found.","Search matched 0 entries"]}}
```

</details>


<details>
<summary>Example 75: Window around a sheet record — executes and exports</summary>

```json
{
  "target": {
    "id": "ws/<sheetId>",
    "beforeLines": 1,
    "afterLines": 4
  },
  "kind": null,
  "include": null,
  "exclude": null,
  "search": null,
  "maxChars": 2000,
  "searchTerm": null,
  "sheetId": null,
  "range": null,
  "offset": null,
  "options": null,
  "summary": "Window around a sheet record"
}
```

Adaptations: Mapped structured example fields to the corresponding public method; replaced placeholder anchors with actual fixture IDs.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-075.json).

Returned/logged output:

```text
{"recordCount":6,"truncated":false,"metadata":{"revision":"p4eqyg","target":{"id":"ws/x523ab","beforeLines":1,"afterLines":4},"kind":{"tokens":["workbook","sheet","table","formula","thread"]},"include":{"tokens":[]},"exclude":{"tokens":[]},"notices":[]}}
```

</details>


<details>
<summary>Example 76: Find cells matching Age (values or formulas) — executes and exports</summary>

```json
{
  "target": null,
  "kind": null,
  "include": null,
  "exclude": null,
  "search": null,
  "maxChars": 3000,
  "searchTerm": "Age",
  "sheetId": null,
  "range": null,
  "offset": 0,
  "options": {
    "matchCase": null,
    "matchEntireCell": null,
    "useRegex": null,
    "matchFormulas": true,
    "ignoreDiacritics": null,
    "maxResults": 20
  },
  "summary": "Find cells matching Age (values or formulas)"
}
```

Adaptations: Mapped structured example fields to the corresponding public method; replaced placeholder anchors with actual fixture IDs.

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-076.json).

Returned/logged output:

```text
{"recordCount":8,"truncated":false,"metadata":{"revision":"7e1ria","target":{"id":"wb/73zm0m","beforeLines":0,"afterLines":0},"kind":{"requested":"workbook,sheet,match","tokens":["workbook","sheet","match"]},"include":{"tokens":[]},"exclude":{"tokens":[]},"notices":[]}}
```

</details>

### API 43 · range.merge

Merge cells in a range into a single region.

<details>
<summary>Example 77: Merge a header row — executes and exports</summary>

```js
sheet.getRange('B1:H1').merge();
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-077.json).

</details>


<details>
<summary>Example 78: Merge across rows — executes and exports</summary>

```js
sheet.getRange('B2:H4').merge(true);
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-078.json).

</details>

### API 44 · range.unmerge

Unmerge any merged blocks intersecting the range.

<details>
<summary>Example 79: Remove merges touching a range — executes and exports</summary>

```js
sheet.getRange('B1:H4').unmerge();
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-079.json).

</details>

### API 45 · defined-names

Defined names and custom functions.

- Use workbook.names.add(name, refersTo) to create a named range (refersTo can be a Range or a string).
- Use workbook.names.getItem(name) / getItemOrNullObject(name) to access existing names.
- Use workbook.names.items to iterate, and namedItem.delete() to remove.
- Use workbook.names.addRange(name, formula, options) and workbook.names.addFunction(...) as legacy helpers.
- Compatibility alias: workbook.definedNames points to workbook.names.
### API 46 · range.format.fill

Set the fill (background) for a range.

- Fill supports theme tokens (accent1..6, lt1/lt2, etc) or hex colors.
- For shades/emphasis, prefer theme token + transform (lighten/darken/opacity) over manually picking multiple hex variants.
- Use the object form for gradients or patterned fills.
- Set fill to null to clear a range's fill.

<details>
<summary>Example 80: Apply a solid fill using a theme token — executes and exports</summary>

```js
const header = sheet.getRange('A1:C1');
header.format.fill = 'accent1';
header.format.font = { color: 'lt1', bold: true };
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-080.json).

</details>


<details>
<summary>Example 81: Apply a theme color transform (LLM-friendly units) — executes and exports</summary>

```js
const header = sheet.getRange('A1:C1');
header.format.fill = {
  type: 'solid',
  color: { type: 'theme', value: 'accent1', transform: { lighten: 15 } },
};
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-081.json).

</details>


<details>
<summary>Example 82: Clear a range's fill — executes and exports</summary>

```js
sheet.getRange('A2:C10').format.fill = null;
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-082.json).

</details>

### API 47 · range.format.fill.color

Set the fill foreground color via the Fill object.

- Use range.format.fill = ... for full fill configs.

<details>
<summary>Example 83: Update fill via fill.color — executes and exports</summary>

```js
const block = sheet.getRange('A1:B10');
block.format.fill.color = '#0055A4';
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-083.json).

</details>

### API 48 · range.format.font

Control font styling on a range.

- Use header.format.font.<prop> = ... for incremental updates.
- Use header.format.font = { ... } or header.format.setFont({ ... }) for patch-style updates.

<details>
<summary>Example 84: Bold header styling — executes and exports</summary>

```js
const header = sheet.getRange('A1:C1');
header.format.font.bold = true;
header.format.font.size = 14;
// Patch-style assignment is supported too:
header.format.font = { color: '#FFFFFF', bold: true };
header.format.setFont({ name: 'Calibri', italic: true });
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-084.json).

</details>

### API 49 · range.format.borders

Apply borders to a range using presets or per-edge config.

<details>
<summary>Example 85: Apply an outside border preset — executes and exports</summary>

```js
const body = sheet.getRange('A2:C10');
body.format.borders = { preset: 'outside', style: 'thin', color: '#DDDDDD' };
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-085.json).

</details>

### API 50 · range.format.numberFormat

Assign Excel number format codes.

- When passing a 2D array, it must match the range size exactly.
- If you want the same code everywhere, passing a single string is simplest.

<details>
<summary>Example 86: Format currency values — executes and exports</summary>

```js
sheet.getRange('B2:B5').format.numberFormat = '"$"#,##0.00';
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-086.json).

</details>


<details>
<summary>Example 87: Assign a per-cell number format matrix — executes and exports</summary>

```js
const r = sheet.getRange('A1:B2');
r.format.numberFormat = [
  ['0', '0.00'],
  ['0', '0.00'],
];
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-087.json).

</details>

### API 51 · range.format.wrapText

Toggle text wrapping for a range.

<details>
<summary>Example 88: Wrap header text — executes and exports</summary>

```js
sheet.getRange('A1:D1').format.wrapText = true;
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-088.json).

</details>

### API 52 · range.format.horizontalAlignment

Set horizontal alignment for a range.

<details>
<summary>Example 89: Center align a header row — executes and exports</summary>

```js
sheet.getRange('A1:D1').format.horizontalAlignment = "center";
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-089.json).

</details>

### API 53 · range.format.autofitColumns

Autofit column widths for a range.

<details>
<summary>Example 90: Autofit columns for the range — executes and exports</summary>

```js
sheet.getRange('A1:D20').format.autofitColumns();
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-090.json).

</details>

### API 54 · range.format.autofitRows

Autofit row heights for a range.

<details>
<summary>Example 91: Autofit rows for the range — executes and exports</summary>

```js
sheet.getRange('A1:D20').format.autofitRows();
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-091.json).

</details>

### API 55 · range.format.rowHeight

Set row height (points) for a range.

- Use rowHeightPx for pixel-based layout sizing.

<details>
<summary>Example 92: Set row height to 18pt — executes and exports</summary>

```js
sheet.getRange('A1:D10').format.rowHeight = 18;
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-092.json).

</details>

### API 56 · range.format.rowHeightPx

Set row height (pixels) for a range.

- Pixels are preferred for layout-sensitive dashboards.

<details>
<summary>Example 93: Set row height to 24px — executes and exports</summary>

```js
sheet.getRange('A1:D10').format.rowHeightPx = 24;
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-093.json).

</details>

### API 57 · range.format.columnWidth

Set column width (Excel width units) for a range.

- Use columnWidthPx for pixel-based layout sizing.

<details>
<summary>Example 94: Set column width to 2.75 — executes and exports</summary>

```js
sheet.getRange('A1:D10').format.columnWidth = 2.75;
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-094.json).

</details>

### API 58 · range.format.columnWidthPx

Set column width (pixels) for a range.

- Pixels are preferred for layout-sensitive dashboards.

<details>
<summary>Example 95: Set column width to 120px — executes and exports</summary>

```js
sheet.getRange('A1:D10').format.columnWidthPx = 120;
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-095.json).

</details>

### API 59 · worksheet.getRange

Get a Range using A1 notation.

- A1 patterns: A1, A1:C10, A:A (entire column), 2:2 (entire row).
- Sheet names with spaces in formulas should be quoted: 'Cost Summary'!A1.

<details>
<summary>Example 96: Create a range from A1 notation — executes and exports</summary>

```js
const sheet = workbook.worksheets.getItem('Sheet1');
const range = sheet.getRange('A1:C10');
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-096.json).

</details>

### API 60 · worksheet.getCell

Get a single cell by zero-based row and column.

- Row and column are 0-based indexes.
- Returns a Range representing a single cell.

<details>
<summary>Example 97: Access a single cell by index — executes and exports</summary>

```js
const sheet = workbook.worksheets.getItem('Sheet1');
const cell = sheet.getCell(1, 4);
cell.values = [['Hello']];
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-097.json).

</details>

### API 61 · range.getCell

Get a single cell inside a range by zero-based offsets.

- Row and column are 0-based offsets from the range's top-left cell.

<details>
<summary>Example 98: Access a cell relative to a range — executes and exports</summary>

```js
const range = sheet.getRange('B2:D4');
const cell = range.getCell(2, 2);
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-098.json).

</details>

### API 62 · range.rowIndex

Return the 0-based row index for the range's top-left cell.

- Throws if the range is not attached to a worksheet.

<details>
<summary>Example 99: Read the range origin row — executes and exports</summary>

```js
const range = sheet.getRange('B2:D4');
range.rowIndex; // 1
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-099.json).

</details>

### API 63 · range.columnIndex

Return the 0-based column index for the range's top-left cell.

- Throws if the range is not attached to a worksheet.

<details>
<summary>Example 100: Read the range origin column — executes and exports</summary>

```js
const range = sheet.getRange('B2:D4');
range.columnIndex; // 1
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-100.json).

</details>

### API 64 · range.getRow

Return a single row inside a range by 0-based offset.

<details>
<summary>Example 101: Select a row inside a block — executes and exports</summary>

```js
const range = sheet.getRange('B2:D4');
const row = range.getRow(1); // B3:D3
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-101.json).

</details>

### API 65 · range.getColumn

Return a single column inside a range by 0-based offset.

<details>
<summary>Example 102: Select a column inside a block — executes and exports</summary>

```js
const range = sheet.getRange('B2:D4');
const col = range.getColumn(0); // B2:B4
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-102.json).

</details>

### API 66 · range.getRangeByIndexes

Return a subrange inside a range using 0-based offsets and sizes.

- Throws if the requested subrange exceeds the current range bounds.

<details>
<summary>Example 103: Select a sub-rectangle — executes and exports</summary>

```js
const range = sheet.getRange('B2:D4');
const sub = range.getRangeByIndexes(1, 1, 2, 2); // C3:D4
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-103.json).

</details>

### API 67 · range.getOffsetRange

Return a range shifted by offsets, optionally changing its size.

<details>
<summary>Example 104: Shift a block down — executes and exports</summary>

```js
const range = sheet.getRange('B2:D4');
const shifted = range.getOffsetRange(1, 0); // B3:D5
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-104.json).

</details>


<details>
<summary>Example 105: Build a small window relative to a block — executes and exports</summary>

```js
const range = sheet.getRange('B2:D4');
const window = range.getOffsetRange(-1, 0, 1, 3); // B1:D1
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-105.json).

</details>

### API 68 · range.getResizedRange

Return a range resized by delta rows/columns, anchored at the current top-left cell.

- Compatibility helper modeled after Office.js Range.getResizedRange(deltaRows, deltaColumns).
- If you have a matrix and just want to write it starting at a single cell, prefer a spill write: sheet.getRange('B2').values = matrix.

<details>
<summary>Example 106: Grow a range by one column — executes and exports</summary>

```js
const range = sheet.getRange('B2:D4');
const wider = range.getResizedRange(0, 1); // B2:E4
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-106.json).

</details>


<details>
<summary>Example 107: Shrink a range by one row and one column — executes and exports</summary>

```js
const range = sheet.getRange('B2:D4');
const smaller = range.getResizedRange(-1, -1); // B2:C3
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-107.json).

</details>

### API 69 · range.getResizeRange

Return a range resized to an explicit row/column count, anchored at the current top-left cell.

- Granola convenience helper (absolute size).
- Equivalent to range.getOffsetRange(0, 0, rowCount, columnCount).

<details>
<summary>Example 108: Resize to a specific shape — executes and exports</summary>

```js
const range = sheet.getRange('B2:D4');
const block = range.getResizeRange(5, 2); // B2:C6
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-108.json).

</details>

### API 70 · range.offset

Alias for range.getOffsetRange(...).

- Use range.getOffsetRange for the full signature and examples.
### API 71 · range.resize

Alias for range.getResizeRange(rowCount, columnCount).

- This is an absolute resize helper.
- Use range.getResizedRange(deltaRows, deltaColumns) when you want Office.js-style delta resize.
### API 72 · range.write

Write a 2D matrix anchored at this range's top-left (auto-sizes the target).

- Prefer this over hand-building A1 end-cells like 'B2:F17'.
- You can pass overwrite: 'error' to prevent overwriting existing content.
- 1D input is treated as a single row (wrap as [[...]] for 2D).

<details>
<summary>Example 109: Write a matrix without computing the end cell — executes and exports</summary>

```js
const anchor = sheet.getRange('B2');
const block = anchor.write([['A','B'],[1,2],[3,4]]);
// block is the resized range (B2:C4)
block.format.autofitColumns();
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-109.json).

</details>

### API 73 · range.values

Read or write cell values using a 2D array.

- Values are string | number | boolean | Date | null.
- Single-cell ranges spill to fit larger matrices.
- Single-row ranges also accept 1D arrays (treated as a row).
- Single-column ranges also accept 1D arrays (treated as a column).
- Granola auto-fixes common transposed vectors for 1-row/1-col ranges (Nx1 vs 1xN).
- Use range.values = null to clear contents.

<details>
<summary>Example 110: Write a values grid — executes and exports</summary>

```js
const range = sheet.getRange('A1:C2');
range.values = [['Date', 'Dog', 'HR'], ['2025-01-02', 'Nova', 58]];
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-110.json).

</details>

### API 74 · range.formulas

Read or write formulas using a 2D string array.

- Formula strings may include or omit the leading '='.
- Single-cell ranges spill to fit larger matrices.
- Single-row and single-column ranges also accept 1D arrays.
- Granola auto-fixes common transposed vectors for 1-row/1-col ranges (Nx1 vs 1xN).

<details>
<summary>Example 111: Write a formula and recalculate — executes and exports</summary>

```js
sheet.getRange('A1:A3').values = [[5], [2], [3]];
sheet.getRange('B1').formulas = [['=SUM(A:A)']];
workbook.recalculate();
const total = sheet.getRange('B1').values[0]?.[0];
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-111.json).

</details>

### API 75 · range.clear

Clear contents and/or formatting for a range.

- applyTo accepts 'contents', 'formats', or 'all'.
- Clearing formats does not change row/column sizing.
- Clearing ranges does not delete charts, shapes, or images.

<details>
<summary>Example 112: Clear contents — executes and exports</summary>

```js
sheet.getRange('A1:C10').clear();
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-112.json).

</details>


<details>
<summary>Example 113: Clear contents and formats — executes and exports</summary>

```js
sheet.getRange('A1:C10').clear({ applyTo: 'all' });
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-113.json).

</details>

### API 76 · ranges

Range addressing and cell reads/writes.

- Create ranges with worksheet.getRange('A1:C10').
- Write with range.values or range.formulas.
- Use range.rowCount/range.columnCount (or range.getRowCount()/getColumnCount()) to inspect range dimensions.
- Use range.formulasR1C1 when it's easier to generate relative formulas (e.g. RC[-1]) and let the API translate them.
- range.values treats strings starting with '=' as formulas; use "'=..." to force literal text.
- range.values and range.formulas require a rectangular matrix matching the range size; single-cell ranges spill to fit larger matrices.
- Clear with range.clear() or range.values = null.
### API 77 · range.fillDown

Copy the first row down, shifting formulas.

- Copies formulas/values from the first row to the rest of the range.
- Relative references are shifted per row.

<details>
<summary>Example 114: Fill down like Excel drag-fill — executes and exports</summary>

```js
const r = sheet.getRange('B2:B10');
r.getCell(0, 0).formulasR1C1 = [['=RC[-1]*2']];
r.fillDown();
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-114.json).

</details>

### API 78 · range.fillRight

Copy the first column right, shifting formulas.

- Copies formulas/values from the first column to the rest of the range.
- Relative references are shifted per column.

<details>
<summary>Example 115: Fill right like Excel drag-fill — executes and exports</summary>

```js
const r = sheet.getRange('B2:F2');
r.getCell(0, 0).formulasR1C1 = [['=RC[-1]*2']];
r.fillRight();
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-115.json).

</details>

### API 79 · range.fillFrom

Fill a larger one-axis range from a source block, repeating values and shifting formulas.

- Destination must extend source in exactly one direction while sharing the other axis.
- Repeated tails are allowed even when the destination size is not an even multiple of the source size.
- Copies literal values exactly and shifts relative formula references; it does not infer numeric/date series.
- Cell formatting and data validation are not propagated.

<details>
<summary>Example 116: Extend a block downward from a source selection — executes and exports</summary>

```js
const seed = sheet.getRange('B2:C3');
const target = sheet.getRange('B2:C8');
target.fillFrom(seed);
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-116.json).

</details>

### API 80 · range.copyFrom

Copy values/formulas from another range (supports 1x1 broadcast).

- When source is 1x1, copyFrom broadcasts to the destination shape.
- When copying formulas, relative references shift across the destination cells.

<details>
<summary>Example 117: Fill a block of formulas using copyFrom — executes and exports</summary>

```js
// Excel-like drag-fill across a whole block (no manual 2D matrix)
const grid = sheet.getRange('B3:AF17');
const seed = sheet.getRange('B3');
seed.formulasR1C1 = [['=RC[-1]*2']];
grid.copyFrom(seed, 'formulas');
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-117.json).

</details>

### API 81 · sparklines

Sparkline groups for compact trend visuals.

- Use worksheet.sparklines.add({ type, targetRange, sourceData }).
- Use worksheet.sparklines.getAll() to list groups.
- Use worksheet.sparklines.deleteAll() to remove all groups (safe no-op when empty).
- worksheet.sparklines is an alias for worksheet.sparklineGroups.
### API 82 · worksheet.sparklines.add

Create a sparkline group for a target range.

- targetRange and sourceData accept a Range or an A1 string address.
- Use worksheet.sparklines.deleteAll() to clear everything.

<details>
<summary>Example 118: Add a line sparkline group — executes and exports</summary>

```js
sheet.getRange('A1:A6').values = [[1],[2],[3],[2],[4],[6]];
const group = sheet.sparklines.add({
  type: 'line',
  sourceData: sheet.getRange('A1:A6'),
  targetRange: sheet.getRange('B1:B6'),
});
group.displayXAxis = true;
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-118.json).

</details>

### API 83 · styles

Formatting ranges (fills, fonts, borders, number formats).

- Use range.format.fill, range.format.font, and range.format.borders.
- Set range.format.numberFormat or range.format.wrapText.
- Use range.format.autofitColumns/Rows for sizing.
### API 84 · formatting-quick-reference

Copy/paste snippets for common range formatting tasks.

- numberFormat accepts either a single string (broadcast) or a 2D string matrix matching the range size.
- Patch-style font updates are supported via range.format.font = { ... } or range.format.setFont({ ... }).
- Bulk updates are supported via range.format = { fill, font, borders, numberFormat, ... }.

<details>
<summary>Example 119: Formatting quick reference — executes and exports</summary>

```js
const range = sheet.getRange('A1:D10');

// Fill + font
range.format.fill = 'lt2';
range.format.font = { name: 'Calibri', size: 11, color: 'tx1' };

// Bulk format (single object assignment)
range.format = {
  fill: 'lt2',
  font: { name: 'Calibri', size: 11, color: 'tx1' },
  borders: { preset: 'outside', style: 'thin', color: '#D1D5DB' },
  horizontalAlignment: 'center',
  verticalAlignment: 'center',
  wrapText: true,
};

// Theme token + transform (prefer this over manual hex shading)
range.getRow(0).format.fill = {
  type: 'solid',
  color: { type: 'theme', value: 'accent1', transform: { darken: 10 } },
};

// Borders (edge + inside presets or per-edge)
range.format.borders = { preset: 'outside', style: 'thin', color: '#D1D5DB' };
range.format.borders.getItem('EdgeBottom').style = 'thin';
range.format.borders.getItem('EdgeBottom').color = '#D1D5DB';

// Alignment + wrapping
range.format.horizontalAlignment = 'center';
range.format.verticalAlignment = 'center';
range.format.wrapText = true;

// Autofit
range.format.autofitColumns();
range.format.autofitRows();

// Number formats
range.format.numberFormat = '0.00';
sheet.getRange('E1:F2').format.numberFormat = [['0', '0.00'], ['0', '0.00']];
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-119.json).

</details>

### API 85 · tables

Excel table objects on a worksheet.

- Use worksheet.tables.add('A1:D10', true) (or worksheet.getRange('A1:D10')) to create a table.
- Use worksheet.tables.getItem(name) to access an existing table.
- Use worksheet.tables.getItemAt(index) to access tables by position.
- Use worksheet.tables.getItemOrNullObject(name) when the table may not exist; check table.isNullObject before using it.
- Use table.delete() to remove a single table, or worksheet.tables.deleteAll() to remove all tables.
### API 86 · data-tables

What-if data tables for formulas.

- Use worksheet.dataTables.add('C2:E5', { rowInput, columnInput }) with the formula in the top-left cell of the range.
- Data tables require at least a rowInput or columnInput.
### API 87 · range.dataValidation

Configure data validation rules for a range.

- You can use either compat shorthand fields (like list: { source }) or Granola's rule/prompt/errorAlert fields.
- Use range.dataValidation = null to clear validation for a range.

<details>
<summary>Example 120: Create a list validation rule — executes and exports</summary>

```js
const range = sheet.getRange('B2:B6');
range.dataValidation = {
  allowBlank: true,
  list: { inCellDropDown: true, source: ['Run', 'Bike', 'Rest'] },
};
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-120.json).

</details>


<details>
<summary>Example 121: Require whole numbers between 1 and 100 — executes and exports</summary>

```js
const range = sheet.getRange('C2:C100');
range.dataValidation = {
  rule: { type: 'whole', operator: 'between', formula1: 1, formula2: 100 },
  errorAlert: {
    style: 'stop',
    title: 'Invalid value',
    message: 'Enter a whole number between 1 and 100.',
  },
};
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-121.json).

</details>


<details>
<summary>Example 122: Constrain dates to a specific year — executes and exports</summary>

```js
const range = sheet.getRange('D2:D200');
range.dataValidation = {
  rule: {
    type: 'date',
    operator: 'between',
    formula1: '=DATE(2026,1,1)',
    formula2: '=DATE(2026,12,31)',
  },
};
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-122.json).

</details>

### API 88 · worksheets

Worksheet collection and selection.

- Use workbook.worksheets.add(name) or add(names) to create sheets.
- Use workbook.worksheets.getItem(name) to select by name.
- Use workbook.worksheets.getItemAt(index) to select by index.
- Use workbook.worksheets.getActiveWorksheet() for the active sheet.
### API 89 · worksheet.freezePanes

Freeze or unfreeze worksheet panes.

<details>
<summary>Example 123: Freeze header row and first two columns — executes and exports</summary>

```js
const sheet = workbook.worksheets.getItem('Data');
sheet.freezePanes.freezeRows(1);
sheet.freezePanes.freezeColumns(2);
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-123.json).

</details>

### API 90 · worksheet.showGridLines

Worksheet gridline visibility surface.
### API 91 · worksheet.getUsedRange

Return the used range for a worksheet.

- Used range is the smallest rectangle covering cells that have been created (values, formulas, or formatting).
- Use sheet.getUsedRange(true) to approximate values-only used bounds.

<details>
<summary>Example 124: Clear the used range — executes and exports</summary>

```js
const sheet = workbook.worksheets.getItem('Sheet1');
const used = sheet.getUsedRange();
used.clear({ applyTo: 'all' });
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-124.json).

</details>

### API 92 · worksheet.reset

Clear the sheet and delete common objects (tables, charts, drawings, sparklines).

- Defaults clear the used range (applyTo='all') and delete tables/charts/drawings/sparklines.
- Pass options to keep objects or to clear formats/contents selectively.

<details>
<summary>Example 125: Reset a worksheet for a clean rebuild — executes and exports</summary>

```js
const sheet = workbook.worksheets.getItem('Sheet1');
sheet.reset();
```

Result: Executed; XLSX checkpoint exported.

[Full receipt](files/runs/api-example-125.json).

</details>


## Exact feature contracts and settings

These are captured contract references, reproduced for completeness. Some conflict with the executed behavior; the corrections in the main guide take precedence as descriptions of this installed version. Links inside each snapshot lead to the separately saved source.

<details>
<summary>api/references/charts-drawings.spec.md</summary>

[Saved source](files/sources/package/api/references/charts-drawings.spec.md).

#### Charts And Drawings

Worksheets expose drawing collections for charts, images, and shapes.

##### Charts

```ts
const chart = sheet.charts.add("bar", {
  from: { row: 1, col: 4 },
  extent: { widthPx: 420, heightPx: 260 },
});
chart.title = "Values";
chart.categories = ["A", "B", "C"];
const series = chart.series.add("Value");
series.values = [3, 5, 8];
series.categories = chart.categories;
```

Charts can be resolved from inspect anchors:

```ts
const chart = workbook.resolve("ch/<chartId>");
```

##### Chart To Image

```ts
const result = await workbook.chartToImage(0, chart.id);
```

##### Images

```ts
const image = sheet.images.add({
  blob: imageBytes,
  contentType: "image/png",
  anchor: {
    from: { row: 1, col: 2 },
    extent: { widthPx: 240, heightPx: 160 },
  },
});

image.bytes;
image.contentType;
image.prompt;
image.uri;
```

##### Shapes

```ts
const shape = sheet.shapes.add({
  geometry: "rect",
  anchor: {
    from: { row: 2, col: 3 },
    extent: { widthPx: 240, heightPx: 80 },
  },
  fill: "#f8fafc",
  line: { fill: "#334155", width: 1 },
});
shape.text = "Label";
```

##### Sparklines

```ts
sheet.sparklineGroups.add({
  targetRange: sheet.getRange("D2:D10"),
  sourceData: sheet.getRange("B2:C10"),
  type: "line",
});
```

`sheet.sparklines` is a deprecated alias for `sheet.sparklineGroups`.

##### Auto Layout

```ts
sheet.autoLayoutDrawings(sheet.charts.items, {
  direction: "vertical",
  frame: { startCell: "E2", width: 800, height: 1200 },
  gap: 16,
  padding: 24,
});
```

##### Cookbook

```ts
sheet.getRange("A1:B4").values = [
  ["Name", "Value"],
  ["A", 3],
  ["B", 5],
  ["C", 8],
];

sheet.charts.add("bar", {
  from: { row: 1, col: 4 },
  extent: { widthPx: 420, heightPx: 260 },
});
```

</details>

<details>
<summary>api/references/comments-notes-names.spec.md</summary>

[Saved source](files/sources/package/api/references/comments-notes-names.spec.md).

#### Comments, Notes, And Names

Workbook comments, notes, and names are root-level facades. Sheet-scoped names
are available from `worksheet.names`.

##### Comments

```ts
workbook.comments.setSelf({
  displayName: "Reviewer",
  initials: "RV",
});

const thread = workbook.comments.addThread(
  { cell: sheet.getRange("B2") },
  "Review this cell",
);

thread.addReply("Updated.");
thread.resolve();
thread.reopen();
```

Threads appear in `workbook.inspect({ kind: "thread" })` and resolve from
`th/...` anchors.

##### Notes

```ts
workbook.notes;
```

Notes serialize through `workbook.toProto()` and participate in recorded
workbook state.

##### Workbook-Scoped Names

```ts
workbook.names.addRange("SalesRange", "Data!$A$1:$A$3", {
  description: "Sales data",
});

const namedRange = workbook.names.getItem("SalesRange");
const range = namedRange.getRange();
namedRange.delete();
```

##### Functions

```ts
workbook.names.addFunction("AddTax", {
  formula: "=LAMBDA(amount, amount * 1.1)",
  description: "Add tax",
});
```

##### Sheet-Scoped Names

```ts
sheet.names.addRange("LocalRange", "Sheet1!$A$1:$A$10");
```

##### Alias

```ts
workbook.definedNames === workbook.names;
```

##### Cookbook

```ts
const data = workbook.worksheets.add("Data");
data.getRange("A1:A3").values = [[1], [2], [3]];
workbook.names.add("SalesRange", data.getRange("A1:A3"));

workbook.comments.setSelf({ displayName: "Reviewer", initials: "RV" });
workbook.comments.addThread({ cell: data.getRange("A1") }, "Check value.");
```

</details>

<details>
<summary>api/references/formatting.spec.md</summary>

[Saved source](files/sources/package/api/references/formatting.spec.md).

#### Formatting, Validation, And Theme

Use range formatting for cell style, worksheet collections for validation and
conditional formatting, and `workbook.setColorScheme` for workbook theme colors.

##### Range Format

```ts
range.format = {
  fill: "#f8fafc",
  font: {
    name: "Aptos",
    bold: true,
    color: "#0f172a",
    size: 11,
  },
  borders: {
    bottom: { style: "thin", color: "#94a3b8" },
  },
  numberFormat: "$#,##0.00",
  horizontalAlignment: "center",
  verticalAlignment: "middle",
  wrapText: true,
  rowHeightPx: 28,
  columnWidthPx: 120,
};
```

##### Number Format

```ts
range.setNumberFormat("0.0%");
```

##### Conditional Formatting

```ts
sheet.getRange("B2:B10").conditionalFormats.add("cellIs", {
  operator: "greaterThan",
  formula: 10,
  format: {
    fill: "#dcfce7",
    font: { bold: true, color: "#166534" },
  },
});
```

Use `workbook.getConditionalFormattingRenderCache(sheet.name)` when a renderer
needs evaluated conditional-format metadata.

##### Data Validation

```ts
sheet.getRange("A2:A10").dataValidation = {
  list: {
    source: ["Open", "Closed"],
    inCellDropDown: true,
  },
  allowBlank: true,
};
```

The range setter accepts Office-style list assignments or validation config
objects.

##### Freeze Panes

```ts
sheet.freezePanes.freezeRows(1);
sheet.freezePanes.freezeColumns(1);
sheet.freezePanes.unfreeze();
```

##### Theme

```ts
workbook.setColorScheme({
  name: "Workbook Theme",
  themeColors: {
    accent1: "#2563eb",
    accent2: "#0f766e",
    bg1: "#ffffff",
    tx1: "#0f172a",
  },
});
```

##### Cookbook

```ts
const header = sheet.getRange("A1:C1");
header.format = {
  fill: "#0f172a",
  font: { bold: true, color: "#ffffff" },
  horizontalAlignment: "center",
};
sheet.freezePanes.freezeRows(1);
```

</details>

<details>
<summary>api/references/formulas.spec.md</summary>

[Saved source](files/sources/package/api/references/formulas.spec.md).

#### Formulas

Workbook formulas are stored on ranges and evaluated by `workbook.recalculate()`.

##### Assign Formulas

```ts
sheet.getRange("A1:B2").values = [
  [2, 3],
  [5, 7],
];

sheet.getRange("C1:C2").formulas = [["=A1+B1"], ["=A2+B2"]];
workbook.recalculate();
```

Strings starting with `=` assigned through `values` are treated as formulas.
Use a leading apostrophe for literal strings:

```ts
sheet.getRange("C1").values = [["=SUM(A1:B1)"]];
sheet.getRange("D1").values = [["'=SUM(A1:B1)"]];
```

##### Dynamic Arrays And Spill

```ts
sheet.getRange("E1").formulas = [["=SEQUENCE(3, 2)"]];
workbook.recalculate();
```

Spill projection cells are not directly editable formula anchors.

##### Trace

```ts
const trace = workbook.trace("Sheet1!C1");
```

`trace` recalculates first and returns a dependency tree for the target cell or
`null` for an invalid reference.

##### Formula Usage Stats

```ts
const stats = workbook.collectFormulaUsageStats();
```

##### Names

```ts
workbook.names.addRange("SalesRange", "Data!$A$1:$A$3", {
  description: "Sales data",
});

workbook.names.addFunction("AddTax", {
  formula: "=LAMBDA(amount, amount * 1.1)",
  description: "Add tax",
});

sheet.names.addRange("LocalRange", "Sheet1!$A$1:$A$10");
```

##### Cookbook

```ts
const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Calc");
sheet.getRange("A1:A3").values = [[1], [2], [3]];
sheet.getRange("B1").formulas = [["=SUM(A1:A3)"]];
workbook.recalculate();

const trace = workbook.trace("Calc!B1");
```

</details>

<details>
<summary>api/references/import-export.md</summary>

[Saved source](files/sources/package/api/references/import-export.md).

#### Import, Export, HTML, Images, And Google Sheets

Workbook import/export helpers cover CSV, Markdown, HTML copy/paste,
image/chart conversion, rendering, and Google Sheets adapter flows.

##### CSV

```ts
const workbook = await Workbook.fromCSV(csv, {
  sheetName: "Data",
});

const { sheet, range } = await workbook.fromCSV(csv, {
  sheetName: "Imported",
});
```

##### Markdown Table

```ts
const workbook = await Workbook.fromMarkdown(markdown, {
  sheetName: "Data",
  format: true,
});
```

Markdown import expects a Markdown table. With `format: true`, the imported
range is formatted as a worksheet table and columns are autofit.

##### HTML

```ts
const html = source.toHTML(0, "A1:B2");
const withFormulas = source.toHTML(0, "A1:B2", { formulas: true });

const result = target.fromHTML(0, html, "C3");
const filled = target.fromHTML(0, html, "B2:C4");
```

`sheetIndex` is zero-based. The optional range uses A1 syntax.

##### Image Import

```ts
const result = workbook.fromImage(
  0,
  { bytes: imageBytes, contentType: "image/png" },
  "A1:D10",
);
```

###### Worksheet Images

```ts
const sheet = workbook.worksheets.getItem("Data");
sheet.images.add({
  svg: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" /></svg>',
  anchor: {
    from: { row: 0, col: 4 },
    extent: { widthPx: 48, heightPx: 48 },
  },
});
```

##### Chart To Image

```ts
const result = await workbook.chartToImage(0, chart.id);
```

##### Render And Export

```ts
const renderBlob = await workbook.render({
  sheet: "Data",
  range: "A1:C10",
  format: "png",
  scale: 2,
});

const exportBlob = await workbook.export({
  sheetName: "Data",
  range: "A1:C10",
  format: "png",
  scale: 2,
  headers: true,
});

const layoutBlob = await workbook.export({
  sheetName: "Data",
  range: "A1:C10",
  format: "layout",
});

const xlsxBlob = await workbook.export({ format: "xlsx" });
```

`format: "layout"` exports JSON evidence for workbook template
reconstruction. It includes sheet/range geometry, row and column metrics,
per-cell frames, style records, style regions, merged ranges, validations,
defined names, and drawing anchors. Use it as inspection evidence; keep authored
workbook creation scripts semantic instead of replaying the layout JSON at
runtime. Layout exports support sheet selection, `range`, and `autoCrop`;
`center`, `width`, and `height` are raster export options.

##### Export Inline Types

```ts
type WorkbookExportFormat = "png" | "jpeg" | "layout" | "xlsx";

type WorkbookExportOptions = {
  fileName?: string;
  sheet?: Worksheet | string | number;
  sheetName?: string;
  sheetIndex?: number;
  range?: string;
  center?: string;
  width?: number;
  height?: number;
  scale?: number;
  autoCrop?: "all" | "charts";
  headers?: boolean;
  format?: WorkbookExportFormat;
  quality?: number;
};
```

##### Google Sheets

```ts
const workbook = await Workbook.fromGoogleSheets(config);
workbook.configureGoogleSheets(config);

const { patch } = workbook.record(() => {
  workbook.worksheets.getItem("Data").getRange("A1").values = [["Updated"]];
});

await workbook.apply(patch, { target: "googleSheets" });
```

##### Cookbook

```ts
const workbook = await Workbook.fromCSV(csv, { sheetName: "Data" });
const preview = await workbook.export({
  sheetName: "Data",
  range: "A1:D20",
  format: "png",
  scale: 2,
});
const proto = workbook.toProto();
```

</details>

<details>
<summary>api/references/inspect-help.md</summary>

[Saved source](files/sources/package/api/references/inspect-help.md).

#### Inspect, Resolve, Find Cells, And Help

Use `inspect`, `resolve`, `findCells`, and `help` to edit imported workbooks
without guessing object identities.

##### Inspect

```ts
const result = await workbook.inspect({
  kind: "sheet,table,formula,chart,thread",
  include: "id,text",
  exclude: "status",
  search: "Revenue",
  maxChars: 8000,
});

console.log(result.ndjson);
```

`inspect` returns bounded NDJSON records and metadata for parsed options,
unknown tokens, search filters, truncation, and target context.

##### Targeted Inspect

```ts
const focused = await workbook.inspect({
  target: { id: sheetAnchorId, beforeLines: 1, afterLines: 6 },
  kind: "sheet,table",
});
```

##### Resolve

```ts
const workbookTarget = workbook.resolve("wb/<id>");
const sheet = workbook.resolve("ws/<sheet-id-or-name>");
const thread = workbook.resolve("th/<thread-id>");
const chart = workbook.resolve("ch/<chart-id>");
```

`resolve` maps inspect anchors to editable facades.

##### Find Cells

```ts
const matches = workbook.findCells({
  searchTerm: "East",
  sheetId: "Revenue",
  options: {
    maxResults: 10,
    matchFormulas: true,
  },
});
```

Result matches include sheet, address, value/formula match metadata, and total
count.

##### Help

```ts
const result = workbook.help("*", {
  search: "setColorScheme|worksheet.charts.add",
  include: ["index", "examples", "notes"],
  maxChars: 12000,
});
console.log(result.ndjson);
```

Common queries:

```text
*, workbook.setColorScheme, worksheet.charts.add, formulas, enum.ChartType
```

##### Create/Edit Loop

```ts
const before = await workbook.inspect({
  kind: "sheet,table,formula,thread,chart",
  search: "Status",
  maxChars: 8000,
});

const sheet = workbook.resolve(sheetAnchorId);
sheet.getRange("B2").values = [["Closed"]];

const after = await workbook.inspect({
  target: { id: sheetAnchorId, beforeLines: 1, afterLines: 6 },
  kind: "sheet,table",
  search: "Closed",
});
```

##### Inline Types

```ts
type WorkbookInspectOptions = {
  fileName?: string | null;
  target?: { id: string; beforeLines?: number; afterLines?: number };
  kind?: string;
  include?: string;
  exclude?: string;
  search?: string;
  maxChars?: number;
};

type WorkbookInspectResult = {
  readonly records: readonly WorkbookInspectRecord[];
  readonly recordCount: number;
  ndjson: string;
  truncated: boolean;
  metadata: WorkbookInspectMetadata;
};

type WorkbookHelpOptions = {
  include?: string[] | string;
  search?: string;
  maxChars?: number;
};
```

</details>

<details>
<summary>api/references/ranges.spec.md</summary>

[Saved source](files/sources/package/api/references/ranges.spec.md).

#### Ranges

Ranges read and write cell values, formulas, formatting, validation,
conditional formats, sparklines, merges, fills, and copies.

##### Addressing

```ts
const range = sheet.getRange("A1:B2");
const byIndexes = sheet.getRangeByIndexes(0, 0, 2, 2);
const cell = sheet.getCell(0, 0);

range.address;
range.getAddress();
range.rowIndex;
range.columnIndex;
range.rowCount;
range.columnCount;
```

##### Values

```ts
sheet.getRange("A1:B2").values = [
  [1, 2],
  [3, 4],
];

sheet.getRange("B1:I1").values = [1, 2, 3, 4, 5, 6, 7, 8];
sheet.getRange("A2:A5").values = ["a", "b", "c", "d"];
```

Single-cell ranges can spill larger matrices:

```ts
sheet.getRange("N9").values = [[1, 2, 3]];
sheet.getRange("B2").values = [[1], [2], [3]];
```

##### Formulas

```ts
sheet.getRange("C1").values = [["=SUM(A1:B1)"]];
sheet.getRange("D1").values = [["'=SUM(A1:B1)"]];

sheet.getRange("C1:C2").formulas = [["=A1+B1"], ["=A2+B2"]];
sheet.getRange("B3:E3").formulas = ["=1+1", "=2+2", "=3+3", "=4+4"];
sheet.getRange("G2:G5").formulas = ["=10", "=11", "=12", "=13"];
```

##### Write

```ts
sheet.getRange("A1").write([
  ["Name", "Value"],
  ["A", 1],
]);

sheet.getRange("C1").write({ formulas: [["=SUM(B2:B10)"]] });
sheet.getRange("E1").write({ values: [["Literal"]] }, { overwrite: "error" });
```

##### Format

```ts
range.format = {
  fill: "#f8fafc",
  font: { bold: true, color: "#0f172a" },
  numberFormat: "$#,##0.00",
  horizontalAlignment: "center",
  verticalAlignment: "middle",
  wrapText: true,
  rowHeightPx: 28,
  columnWidthPx: 120,
};

range.setNumberFormat("0.0%");
```

##### Navigation

```ts
range.getCell(0, 1);
range.getRow(0);
range.getColumn(1);
range.getRange("A1");
range.getRangeByIndexes(0, 0, 1, 1);
range.getOffsetRange(1, 0);
range.getResizedRange(2, 3);
range.getResizeRange(10, 4);
range.offset(1, 0);
range.resize(10, 4);
range.getCurrentRegion();
```

##### Clear, Merge, Fill, Copy

```ts
range.clear({ applyTo: "contents" });
range.clear({ applyTo: "formats" });
range.clear({ applyTo: "all" });

range.merge();
range.unmerge();

range.fillDown();
range.fillRight();
range.fillFrom(sourceRange);

range.copyFrom(sourceRange);
range.copyTo(targetRange);
```

##### Inline Types

```ts
type CellValue = string | number | boolean | Date | null;

type RangeWriteOptions = {
  clear?: "contents" | "formats" | "all";
  overwrite?: "allow" | "error";
  resize?: "auto" | "none";
};

type RangeWritePayload =
  | CellValue[][]
  | CellValue[]
  | null
  | {
      values?: CellValue[][] | CellValue[] | null;
      formulas?: string[][] | string[];
      formulasR1C1?: string[][] | string[];
    };
```

##### Cookbook

```ts
const data = sheet.getRange("A1:C4");
data.values = [
  ["Region", "Q1", "Q2"],
  ["North", 10, 12],
  ["South", 8, 9],
  ["Total", null, null],
];
sheet.getRange("B4:C4").formulas = [["=SUM(B2:B3)", "=SUM(C2:C3)"]];
data.format = { numberFormat: "General", wrapText: true };
```

</details>

<details>
<summary>api/references/tables.spec.md</summary>

[Saved source](files/sources/package/api/references/tables.spec.md).

#### Tables

Worksheet tables are created from an existing range and can edit rows, headers,
data body ranges, names, style, and totals.

##### Add Table

```ts
const sheet = workbook.worksheets.add("Inventory");
sheet.getRange("A1:C3").values = [
  ["SKU", "Color", "Stock"],
  ["100-001", "Black", 15],
  ["100-002", "Sage", 40],
];

const table = sheet.tables.add("A1:C3", true, "InventoryTable");
```

##### Read

```ts
sheet.tables.items;
sheet.tables.getItem("InventoryTable");
sheet.tables.getItemAt(0);

table.name;
table.address;
table.getRange();
table.getHeaderRowRange();
table.getDataRows();
```

##### Add Rows

```ts
table.rows.add(null, [["100-003", "Denim", 12]]);
```

##### Headerless Table

```ts
sheet.getRange("A1:B3").values = [
  ["West", 12],
  ["East", 9],
  ["South", 15],
];
sheet.tables.add("A1:B3", false, "HeaderlessTable");
```

##### Recorded Table Edit

```ts
const { patch } = workbook.record(() => {
  const table = sheet.tables.add("A1:C3", true, "InventoryTable");
  table.rows.add(null, [["100-003", "Denim", 12]]);
});
```

The recorded patch uses `table.add` followed by row/column/table operations for
subsequent edits.

##### Cookbook

```ts
const table = sheet.tables.add("A1:D5", true, "DataTable");
table.rows.add(null, [["West", "Ava", 12, 14]]);

sheet.getRange(table.address).format = {
  font: { name: "Aptos" },
  wrapText: true,
};
```

</details>

<details>
<summary>api/references/workbook.spec.md</summary>

[Saved source](files/sources/package/api/references/workbook.spec.md).

#### Workbook Facade

##### Create And Load

```ts
const workbook = Workbook.create();
const [sheet1, sheet2] = workbook.worksheets.add(["Sheet1", "Sheet2"]);
const loaded = Workbook.load(proto);
const validated = Workbook.load(proto, { validate: true });
```

##### Import

```ts
const fromCsv = await Workbook.fromCSV(csv, { sheetName: "Data" });
const fromMarkdown = await Workbook.fromMarkdown(markdown, {
  sheetName: "Data",
  format: true,
});

const { sheet, range } = await workbook.fromCSV(csv, { sheetName: "Data" });
```

##### Root Collections

```ts
workbook.worksheets;
workbook.sheets; // alias
workbook.pivotTables;
workbook.slicers;
workbook.comments;
workbook.notes;
workbook.names;
workbook.definedNames; // alias
workbook.awareness;
```

##### Theme

```ts
workbook.setColorScheme({
  name: "Workbook Theme",
  themeColors: {
    accent1: "#2563eb",
    bg1: "#ffffff",
    tx1: "#0f172a",
  },
});
const theme = workbook.theme;
```

##### Recalculate And Trace

```ts
workbook.recalculate();
const trace = workbook.trace("Data!C2");
const stats = workbook.collectFormulaUsageStats();
```

##### Inspect, Search, Help, Resolve

```ts
const snapshot = await workbook.inspect({ kind: "sheet,table,formula" });
const matches = workbook.findCells({
  searchTerm: "East",
  sheetId: "Revenue",
  options: { maxResults: 10, matchFormulas: true },
});
const help = workbook.help("*", {
  search: "setColorScheme|worksheet.charts.add",
  include: ["index", "examples", "notes"],
});
const target = workbook.resolve(anchorId);
```

##### Record And Apply

```ts
const { result, patch, idMap, crdtUpdateV2 } = workbook.record(() => {
  const sheet = workbook.worksheets.getItem("Data");
  sheet.getRange("A1").values = [["Value"]];
  return sheet;
});

const applied = workbook.apply(patch);
```

##### CRDT

```ts
workbook.hydrateCrdtFromProto();
const ready = workbook.isCollaborativeStateReady();
const unsubscribe = workbook.onCrdtUpdateV2((update, origin) => {
  send(update);
});
workbook.applyCrdtUpdateV2(remoteUpdate, { recalculate: true });
unsubscribe();
```

##### Serialize

```ts
const proto = workbook.toProto();
```

##### Utilities

```ts
workbook.utils.columnToLetter(1); // "A"
workbook.utils.letterToColumn("AA"); // 27
workbook.utils.toA1String(1, 1, 3, 2); // "A1:B3"
workbook.utils.fillRight([[1]], 3); // [[1, 1, 1]]
workbook.utils.fillDown([[1]], 3); // [[1], [1], [1]]
```

##### Cookbook

```ts
const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Data");
sheet.getRange("A1:B2").values = [
  ["Name", "Value"],
  ["A", 1],
];

const { patch } = workbook.record(() => {
  sheet.getRange("B2").values = [[2]];
});
```

</details>

<details>
<summary>api/references/worksheets.spec.md</summary>

[Saved source](files/sources/package/api/references/worksheets.spec.md).

#### Worksheets

`workbook.worksheets` and `workbook.sheets` expose the worksheet collection.

##### Collection

```ts
const sheet = workbook.worksheets.add("Data");
const existing = workbook.worksheets.getItem("Data");
const maybe = workbook.worksheets.getItemOrNullObject("Data");
const first = workbook.worksheets.getFirst();
const active = workbook.worksheets.getActiveWorksheet();

workbook.worksheets.setActiveWorksheet("Data");
const count = workbook.worksheets.getSheetCount();
const index = workbook.worksheets.getSheetIndex("Data");
const name = workbook.worksheets.getSheetNameByIndex(0);
const byIndex = workbook.worksheets.getItemAt(0);
```

##### Worksheet Properties

```ts
sheet.name = "Renamed";
sheet.index = 0;
sheet.showGridLines = false;
sheet.tabColor = "#1F4E78";

sheet.id;
sheet.sheetId;
sheet.tabColor;
sheet.defaultRowHeight;
sheet.defaultColWidth;
sheet.baseColWidth;
```

##### Ranges

```ts
const range = sheet.getRange("A1:B2");
const byIndexes = sheet.getRangeByIndexes(0, 0, 2, 2);
const cell = sheet.getCell(0, 0);
const used = sheet.getUsedRange();
const usedValuesOnly = sheet.getUsedRange(true);
```

`getRange` accepts A1 strings only. Use zero-based row/column indexes with
`getRangeByIndexes` and `getCell`.

##### Worksheet Collections

```ts
sheet.tables;
sheet.charts;
sheet.shapes;
sheet.images;
sheet.pivotTables;
sheet.slicers;
sheet.sparklineGroups;
sheet.conditionalFormattings;
sheet.dataValidations;
sheet.freezePanes;
sheet.names;
sheet.cells;
```

##### Reset And Delete

```ts
sheet.reset({
  clear: "used",
  applyTo: "all",
  deleteTables: true,
  deleteCharts: true,
  deleteDrawings: true,
  deleteSparklines: true,
});

sheet.delete();
```

##### Merge

```ts
sheet.mergeCells("A1:B2");
sheet.unmergeCells("A1:B2");
```

##### Drawings

```ts
sheet.deleteAllDrawings();
sheet.autoLayoutDrawings(items, options);
```

##### Inline Types

```ts
type WorksheetResetOptions = {
  clear?: "used" | "none";
  applyTo?: "contents" | "formats" | "all";
  deleteTables?: boolean;
  deleteCharts?: boolean;
  deleteDrawings?: boolean;
  deleteSparklines?: boolean;
};
```

##### Cookbook

```ts
const sheet = workbook.worksheets.add("Data");
sheet.getRange("A1:B2").values = [
  ["Name", "Value"],
  ["A", 1],
];
sheet.mergeCells("D1:E1");
sheet.freezePanes.freezeRows(1);
```

</details>

<details>
<summary>references/comments.md</summary>

[Saved source](files/sources/package/references/comments.md).

#### Comments

Workbook comments use people, threads, replies, reactions, and thread state.

##### Current Author

```ts
workbook.comments.setSelf({
  displayName,
  initials,
  email,
});
```

##### Person Inline Type

```ts
type PersonConfig = {
  id?: string;
  displayName: string;
  initials?: string;
  email?: string;
  avatarUrl?: string;
};
```

##### Cell And Range Threads

```ts
const cellThread = workbook.comments.addThread(
  { cell: sheet.getRange(cellAddress) },
  bodyText,
);

const rangeThread = workbook.comments.addThread(
  { range: sheet.getRange(rangeAddress) },
  bodyText,
);
```

##### Thread Inline Types

```ts
type WorkbookCommentTarget =
  | { cell: Range }
  | { range: Range };

type ThreadAddOptions = {
  author?: Person | PersonConfig | { id: string };
  createdAt?: string;
  position?: { x: number; y: number; unit?: "px" | "emu" };
};
```

##### Replies And State

```ts
const reply = cellThread.addReply(replyText);
reply.toggleReaction(reactionText);
cellThread.resolve();
cellThread.reopen();
```

</details>

<details>
<summary>references/conditional-formatting.spec.md</summary>

[Saved source](files/sources/package/references/conditional-formatting.spec.md).

#### Conditional Formatting

Use `range.conditionalFormats` to add formatting rules to ranges.

##### Rule Pattern

```ts
range.conditionalFormats.add(ruleType, {
  operator,
  formula,
  format,
});
```

Choose `ruleType`, `operator`, color, and style strings from the inline types below.

##### Rule Inline Types

```ts
type ConditionalFormatRuleType =
  | "cellIs" | "CellValue" | "Custom" | "expression"
  | "colorScale" | "dataBar" | "iconSet"
  | "containsText" | "notContainsText" | "beginsWith" | "endsWith"
  | "containsBlanks" | "notContainsBlanks" | "containsErrors" | "notContainsErrors"
  | "duplicateValues" | "uniqueValues" | "timePeriod" | "top10" | "aboveAverage";

type CellIsOperator =
  | "greaterThan"
  | "greaterThanOrEqual"
  | "lessThan"
  | "lessThanOrEqual"
  | "equal"
  | "notEqual"
  | "between"
  | "notBetween";

type ConditionalFormatConfig =
  | { operator: CellIsOperator; formula: string | number | Array<string | number>; format?: DifferentialFormatConfig }
  | { formula: string | number; format?: DifferentialFormatConfig }
  | { colors?: ColorConfig[]; thresholds?: CfvoInput[] }
  | { color?: ColorConfig; thresholds?: CfvoInput[]; gradient?: boolean }
  | { iconSet: string; showValue?: boolean; reverse?: boolean; thresholds?: CfvoInput[] }
  | { text: string; format?: DifferentialFormatConfig }
  | { timePeriod: "yesterday" | "today" | "tomorrow" | "last7Days" | "lastWeek" | "thisWeek" | "nextWeek" | "lastMonth" | "thisMonth" | "nextMonth"; format?: DifferentialFormatConfig }
  | { rank?: number; percent?: boolean; bottom?: boolean; format?: DifferentialFormatConfig }
  | { aboveAverage?: boolean; equalAverage?: boolean; stdDev?: number; format?: DifferentialFormatConfig };

type DifferentialFormatConfig = {
  fill?: FillConfig;
  font?: { bold?: boolean; italic?: boolean; color?: ColorConfig };
  border?: RangeBordersConfig;
  numberFormat?: string;
};

type CfvoInput =
  | "min"
  | "max"
  | number
  | `${number}%`
  | { type: "min" | "max" | "num" | "percent" | "percentile"; value?: string | number };
```

##### Custom Rule

```ts
range.conditionalFormats.addCustom(expression, {
  fill,
  font,
  border,
});
```

##### Rule Families

Supported rule families include cell value, custom expression, text, blank/error, duplicate/unique, time period, top/bottom, above/below average, color scale, data bar, and icon set.

##### Clear Rules

```ts
range.conditionalFormats.deleteAll();
```

</details>

<details>
<summary>references/data-tables.spec.md</summary>

[Saved source](files/sources/package/references/data-tables.spec.md).

##### Data tables (`worksheet.dataTables`)

Use data tables to project a formula across a row/column grid.

###### Two-variable data table

```ts
const sheet = workbook.worksheets.add("TwoVariable");

sheet.getRange("B3").values = [[0.095]];  // rate
sheet.getRange("B4").values = [[360]];    // periods
sheet.getRange("B5").values = [[80000]];  // principal

sheet.getRange("C2").formulas = [[`=PMT(B3/12,B4,-B5)`]];
sheet.getRange("C3:C5").values = [[0.09], [0.0925], [0.095]];
sheet.getRange("D2:E2").values = [[180, 360]];

sheet.dataTables.add("D3:E5", {
  rowInput: "B4",
  columnInput: "B3",
});

workbook.recalculate();
```

###### Single-variable (row input only)

```ts
sheet.dataTables.add("D3:E3", { rowInput: "B4" });
workbook.recalculate();
```

</details>

<details>
<summary>references/data-validations.spec.md</summary>

[Saved source](files/sources/package/references/data-validations.spec.md).

#### Data Validations

Use range-level or worksheet-level data validation config.

##### Range Assignment

```ts
range.dataValidation = validationConfig;
```

##### Rule Config

```ts
range.dataValidation.rule = {
  type,
  formula1,
  formula2,
  values,
  operator,
};
range.dataValidation.prompt = promptConfig;
range.dataValidation.errorAlert = errorAlertConfig;
range.dataValidation.ignoreBlanks = ignoreBlanks;
range.dataValidation.inCellDropDown = inCellDropDown;
```

##### Validation Inline Types

```ts
type DataValidationRule = {
  type: "none" | "whole" | "decimal" | "list" | "date" | "time" | "textLength" | "custom";
  operator?:
    | "between"
    | "notBetween"
    | "equal"
    | "notEqual"
    | "lessThan"
    | "lessThanOrEqual"
    | "greaterThan"
    | "greaterThanOrEqual";
  formula1?: string | number;
  formula2?: string | number;
  values?: string[];
};

type ValidationConfig = {
  rule?: DataValidationRule | null;
  prompt?: { title?: string; message?: string; show?: boolean } | null;
  errorAlert?: {
    title?: string;
    message?: string;
    style?: "stop" | "warning" | "information";
    show?: boolean;
  } | null;
  ignoreBlanks?: boolean;
  inCellDropDown?: boolean;
};
```

##### Collection Add

```ts
sheet.dataValidations.add({
  range,
  rule,
  prompt,
  errorAlert,
  ignoreBlanks,
  inCellDropDown,
});
```

Validation type, operator, and error style values are the strings listed above.

</details>

<details>
<summary>references/defined-names.spec.md</summary>

[Saved source](files/sources/package/references/defined-names.spec.md).

##### Defined names (`workbook.names`, `worksheet.names`)

Use defined names for reusable ranges and workbook functions.

###### Named ranges

Workbook-scoped named range:

```ts
const data = workbook.worksheets.add("Data");
data.getRange("A1:A3").values = [[1], [2], [3]];

workbook.names.addRange("SalesRange", "Data!$A$1:$A$3", {
  description: "Input sales data",
});

data.getRange("B1").formulas = [["=SUM(SalesRange)"]];
workbook.recalculate();
```

Sheet-scoped named range:

```ts
const sheet1 = workbook.worksheets.add("Sheet1");
sheet1.names.addRange("LocalRange", "Sheet1!$A$1:$A$10");

sheet1.getRange("B1").formulas = [["=SUM(LocalRange)"]];
workbook.recalculate();
```

###### Named functions (LAMBDA)

```ts
workbook.names.addFunction("AddTax", {
  lambda: "LAMBDA(price, tax, price*(1+tax))", // export may auto-prefix "="
  description: "Adds tax to a price",
  parameters: [
    { name: "price", description: "Base price" },
    { name: "tax", description: "Tax rate (e.g. 0.1)" },
  ],
  returns: "Taxed price",
});
```

</details>

<details>
<summary>references/drawings.spec.md</summary>

[Saved source](files/sources/package/references/drawings.spec.md).

#### Worksheet Drawings

Worksheets support charts and shapes anchored to the cell grid.

##### Units

- Anchor row and column indexes are zero-based.
- Offsets and extents use pixels.
- Use `rowOffsetPx`, `colOffsetPx`, `widthPx`, and `heightPx`.

##### Chart Config

```ts
const chart = sheet.charts.add(chartType, {
  from,
  to,
  extent,
  title,
  titleTextStyle,
  categories,
  series,
  hasLegend,
  legend,
  xAxis,
  yAxis,
  dataLabels,
  dataTable,
});
```

##### Chart Inline Types

```ts
type WorksheetChartType =
  | "bar" | "line" | "area" | "pie" | "doughnut" | "scatter"
  | "bubble" | "radar" | "stock" | "treemap" | "sunburst" | "histogram"
  | "boxWhisker" | "waterfall" | "funnel" | "map";

type WorksheetChartConfig = {
  from?: { row: number; col: number; rowOffsetPx?: number; colOffsetPx?: number };
  to?: { row: number; col: number; rowOffsetPx?: number; colOffsetPx?: number };
  extent?: { widthPx?: number; heightPx?: number };
  title?: string;
  titleTextStyle?: TextStyleConfig;
  categories?: string[];
  series?: Array<{ name: string; values?: number[]; categories?: string[]; fill?: FillConfig; line?: LineConfig }>;
  hasLegend?: boolean;
  legend?: { position?: "left" | "top" | "topRight" | "right" | "bottom"; overlay?: boolean; textStyle?: TextStyleConfig };
  xAxis?: ChartAxisConfig;
  yAxis?: ChartAxisConfig;
  dataLabels?: { showValue?: boolean; position?: "center" | "inEnd" | "outEnd"; textStyle?: TextStyleConfig };
  dataTable?: { visible?: boolean; showLegendKey?: boolean; textStyle?: TextStyleConfig };
};

type TextStyleConfig = {
  bold?: boolean;
  italic?: boolean;
  fontSize?: number;
  color?: ColorConfig;
  fill?: FillConfig;
  alignment?: "left" | "center" | "right" | "justify";
};

type ChartAxisConfig = {
  title?: string | { text?: string; textStyle?: TextStyleConfig };
  min?: number;
  max?: number;
  numberFormatCode?: string;
  position?: "left" | "right" | "top" | "bottom";
  tickLabelPosition?: "high" | "low" | "nextTo" | string;
  textStyle?: TextStyleConfig;
  line?: LineConfig;
  majorGridlines?: LineConfig | null;
  minorGridlines?: LineConfig | null;
};
```

##### Chart From Range

```ts
const chart = sheet.charts.add(chartType, sourceRange, seriesBy);
chart.setPosition(startRange, endRange);
chart.width = widthPx;
chart.height = heightPx;
```

##### Shape Config

```ts
const shape = sheet.shapes.add({
  geometry,
  anchor,
  fill,
  line,
  name,
});
```

##### Shape Inline Type

```ts
type WorksheetShapeConfig = {
  geometry: string; // common: "rect", "roundRect", "ellipse", "textbox"; full list is the slide shape preset list
  anchor?: DrawingAnchorConfig;
  fill?: FillConfig;
  line?: LineConfig;
  name?: string;
};

type DrawingAnchorConfig = {
  from: { row: number; col: number; rowOffsetPx?: number; colOffsetPx?: number };
  to?: { row: number; col: number; rowOffsetPx?: number; colOffsetPx?: number };
  extent?: { widthPx?: number; heightPx?: number };
};
```

##### Rebuild And Layout

```ts
sheet.deleteAllDrawings();

const drawings = [...sheet.charts.items, ...sheet.shapes.items];
sheet.autoLayoutDrawings(drawings, {
  direction,
  frame,
  gap,
  padding,
});
```

##### Auto Layout Inline Type

```ts
type WorksheetDrawingAutoLayoutOptions = {
  direction?: "vertical" | "horizontal" | "grid";
  align?: "start" | "center" | "end";
  columns?: number; // used for grid
  gap?: number;
  padding?: number;
  frame?:
    | { left: number; top: number; width: number; height: number }
    | { startCell: string; width: number; height: number };
};
```

</details>

<details>
<summary>references/images.spec.md</summary>

[Saved source](files/sources/package/references/images.spec.md).

#### Worksheet Images

Use `sheet.images.add(...)` to attach images to cell-grid anchors.

##### Add Image

```ts
const image = sheet.images.add({
  path,
  blob,
  dataUrl,
  uri,
  prompt,
  alt,
  anchor,
});

const icon = sheet.images.add({
  svg: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" fill="#2563eb"/></svg>',
  anchor: { from: { row: 1, col: 5 }, extent: { widthPx: 48, heightPx: 48 } },
});
```

##### Image Inline Types

```ts
type WorksheetImageConfig = ImageSource & {
  alt?: string;
  contentType?: string;
  anchor?: DrawingAnchorConfig;
};

type ImageSource =
  | { path: string }
  | { svg: string }
  | { blob: ArrayBuffer }
  | { dataUrl: string }
  | { uri: string }
  | { prompt: string };
```

##### Anchor

```ts
const anchor = {
  from,
  to,
  extent,
};
```

`from`, `to`, and `extent` use the same zero-based anchor and pixel extent shapes as worksheet charts and shapes.

##### Anchor Inline Type

```ts
type DrawingAnchorConfig = {
  from: AnchorInput;
  to?: AnchorInput;
  extent?: {
    widthPx?: number;
    heightPx?: number;
    widthEmu?: number;
    heightEmu?: number;
  };
};

type AnchorInput = {
  row: number;
  col: number;
  rowOffsetPx?: number;
  colOffsetPx?: number;
  rowOffsetEmu?: number;
  colOffsetEmu?: number;
};
```

##### Edit

```ts
image.anchor = anchor;
image.alt = altText;
image.replace(sourceConfig);
image.delete();
```

</details>

<details>
<summary>references/ranges.spec.md</summary>

[Saved source](files/sources/package/references/ranges.spec.md).

##### Ranges API

Use `Worksheet.getRange(address)` to read and write grid data using Excel A1 notation.

###### Values

`range.values` is a 2D array of `CellValue`:

- `string | number | boolean | Date | null`
- `null` represents an empty cell

```ts
sheet.getRange("A1:C1").values = [["Date", "Dog", "Resting HR (bpm)"]];
sheet.getRange("A2:C4").values = [
  ["2025-01-02", "Nova", 58],
  ["2025-01-03", "Nova", 61],
  ["2025-01-03", "Comet", 72],
];
```

###### Checkboxes

Assign boolean values to create Excel checkbox cells. Use `range.values = null` to clear a range and remove the checkbox control.

```ts
sheet.getRange("A1").values = [[true]];
sheet.getRange("A2").values = [[false]];

sheet.getRange("A1:A2").values = null;
```

###### Merged cells

Use `range.merge()` to merge a rectangular range into a single cell region, and `range.unmerge()` to remove merges that intersect the range.

```ts
// Merge a header across columns
sheet.getRange("B1:H1").merge();

// Merge across rows (one merge per row, like Excel “Merge Across”)
sheet.getRange("B2:H4").merge(true);

// Remove merges touching the range
sheet.getRange("B1:H4").unmerge();
```

###### Formulas

`range.formulas` is a 2D array of strings. Call `workbook.recalculate()` before reading computed `values`.

```ts
sheet.getRange("E1").values = [["Average HR"]];
sheet.getRange("E2").formulas = [[`=AVERAGE(C2:C4)`]];

workbook.recalculate();
const average = sheet.getRange("E2").values[0]?.[0];
```

###### A1 addressing conventions

Supported addressing patterns include:

- `A1` (single cell)
- `A1:C10` (rectangle)
- `A:A` (entire column)
- `2:2` (entire row)

In formulas, quote sheet names that contain spaces/punctuation: `'Cost Summary'!A1`.

</details>

<details>
<summary>references/sparklines.spec.md</summary>

[Saved source](files/sources/package/references/sparklines.spec.md).

#### Sparklines

Use `sheet.sparklineGroups` for worksheet sparkline groups.

##### Add Group

```ts
const group = sheet.sparklineGroups.add({
  type,
  targetRange,
  sourceData,
  dateAxisRange,
  seriesColor,
  negativeColor,
  markers,
  axis,
  lineWeight,
  displayEmptyCellsAs,
  displayHidden,
});
```

Sparkline type is a string. Empty-cell display mode and axis min/max modes are proto enum numbers on the current facade; inspect nearby tests before setting them directly.

##### Sparkline Inline Type

```ts
type SparklineConfig = {
  type: "line" | "column" | "stacked";
  targetRange: Range | string;
  sourceData: Range | string;
  dateAxisRange?: Range | string;
  lineWeight?: number;
  displayHidden?: boolean;
  seriesColor?: ColorConfig;
  negativeColor?: ColorConfig;
  axisColor?: ColorConfig;
  markersColor?: ColorConfig;
  firstMarkerColor?: ColorConfig;
  lastMarkerColor?: ColorConfig;
  highMarkerColor?: ColorConfig;
  lowMarkerColor?: ColorConfig;
  markers?: SparklineMarkersOptions;
  axis?: SparklineAxisOptions;
};

type SparklineMarkersOptions = {
  show?: boolean;
  high?: boolean;
  low?: boolean;
  first?: boolean;
  last?: boolean;
  negative?: boolean;
};

type SparklineAxisOptions = {
  showAxis?: boolean;
  manualMin?: number;
  manualMax?: number;
  rightToLeft?: boolean;
};
```

##### Range Alias

```ts
const group = targetRange.sparklines.add(type, sourceRange, sparklineConfig);
```

##### Edit And Delete

```ts
group.seriesColor = colorConfig;
group.markers = markerConfig;
group.axis = axisConfig;
group.delete();
sheet.sparklineGroups.deleteAll();
```

</details>

<details>
<summary>references/styles.spec.md</summary>

[Saved source](files/sources/package/references/styles.spec.md).

#### Range Formatting

Use `range.format` to style cells with fills, fonts, borders, number formats, sizing, and alignment.

##### Shared Color And Fill Inline Types

```ts
type ThemeColorName =
  | "accent1" | "accent2" | "accent3" | "accent4" | "accent5" | "accent6"
  | "bg1" | "bg2" | "tx1" | "tx2" | "dk1" | "lt1" | "dk2" | "lt2"
  | "hlink" | "folHlink";

type ColorConfig =
  | string
  | { type: "rgb"; value: string; transform?: { opacity?: number; lighten?: number; darken?: number } }
  | { type: "theme"; value: ThemeColorName; transform?: { opacity?: number; lighten?: number; darken?: number } };

type FillConfig =
  | string
  | { type: "none" }
  | { type: "solid"; color: ColorConfig }
  | { type: "gradient"; stops: Array<{ offset: number; color: ColorConfig }>; angleDeg?: number; gradientKind?: "linear" | "path" };
```

##### Grouped Format

```ts
range.format = {
  fill,
  font,
  horizontalAlignment,
  verticalAlignment,
  wrapText,
  borders,
  numberFormat,
};
```

##### Format Inline Types

```ts
type RangeFormatConfig = {
  fill?: FillConfig | null;
  font?: RangeFontConfig | null;
  borders?: RangeBordersConfig;
  numberFormat?: string | string[][];
  wrapText?: boolean;
  horizontalAlignment?:
    | "general"
    | "left"
    | "center"
    | "right"
    | "fill"
    | "justify"
    | "centerAcrossSelection"
    | "distributed";
  verticalAlignment?: "top" | "middle" | "bottom" | "center";
  rowHeight?: number; // points
  rowHeightPx?: number;
  columnWidth?: number; // Excel width units
  columnWidthPx?: number;
};

type RangeFontConfig = {
  bold?: boolean;
  italic?: boolean;
  size?: number;
  name?: string;
  color?: ColorConfig;
};
```

##### Sizing

```ts
range.format.rowHeightPx = rowHeightPx;
range.format.columnWidthPx = columnWidthPx;
```

Pixel sizing is the preferred surface for layout-sensitive work. Point row heights and Excel-width column widths are also available through `rowHeight` and `columnWidth`.

##### Borders

```ts
range.format.borders = {
  preset,
  style,
  color,
};
```

Use the border preset strings below. Border style strings are Excel/OpenXML-style names such as `"thin"`, `"medium"`, `"thick"`, `"dashed"`, and `"dotted"`.

##### Border Inline Type

```ts
type RangeBordersConfig =
  | {
      preset: "none" | "outside" | "inside" | "all" | "doubleBottom";
      style?: string; // common: "thin", "medium", "thick", "dashed", "dotted"
      color?: ColorConfig;
    }
  | {
      top?: BorderLineInput;
      bottom?: BorderLineInput;
      left?: BorderLineInput;
      right?: BorderLineInput;
      inside?: BorderLineInput;
      insideHorizontal?: BorderLineInput;
      insideVertical?: BorderLineInput;
      diagonalUp?: BorderLineInput;
      diagonalDown?: BorderLineInput;
    };

type BorderLineInput = {
  style?: string;
  color?: ColorConfig;
  weight?: number;
};
```

</details>

<details>
<summary>references/tables.spec.md</summary>

[Saved source](files/sources/package/references/tables.spec.md).

#### Worksheet Tables

Use `sheet.tables` to add and edit Excel-style tables.

##### Add Table

```ts
const table = sheet.tables.add(rangeAddress, hasHeaders, tableName);
```

##### Add Inline Type

```ts
type WorksheetTableAddArgs = {
  range: string | Range;
  hasHeaders?: boolean;
  name?: string;
};
```

##### Rows And Values

```ts
table.rows.add(rowIndex, rowValues);
table.getRange().values = tableValues;
```

##### Table Style

```ts
table.name = tableName;
table.showTotals = showTotals;
table.style = tableStyleName;
table.styleOptions = tableStyleOptions;
```

##### Table Inline Type

```ts
type WorkbookTableStyleOptions = {
  showHeaders?: boolean;
  showTotals?: boolean;
  bandedRows?: boolean;
  bandedColumns?: boolean;
  firstColumn?: boolean;
  lastColumn?: boolean;
};
```

##### Delete

```ts
table.delete();
sheet.tables.deleteAll();
```

</details>

<details>
<summary>references/workbook.spec.md</summary>

[Saved source](files/sources/package/references/workbook.spec.md).

##### Workbook API

Use `Workbook` to create, edit, recalculate, and export spreadsheet artifacts.

###### Quick start

```ts
import { Workbook } from "@oai/artifact-tool";

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Sheet1");

sheet.getRange("A1").values = [["Hello"]];
workbook.recalculate();

const proto = workbook.toProto();
```

###### Workbook lifecycle

- `Workbook.create()` — start a new workbook.
- `new Workbook(proto)` — hydrate from an existing artifact, edit, then `toProto()` again.
- `workbook.recalculate()` — evaluate formulas.
- `workbook.toProto()` — produce the protobuf-like JSON artifact.

###### Worksheets

- `workbook.worksheets.add(name)` — add (or return) a worksheet by name.
- `workbook.worksheets.getItem(name | index)` — fetch by name or index.
- `workbook.worksheets.getActiveWorksheet()` — fetch the active worksheet.
- `workbook.worksheets.count` — number of sheets.

</details>


## Enums: every listed token

Chart types have the native/export matrix above. All five line styles were distributed across the geometry contact sheets. Every shape enum token was attempted: 189 built-ins succeeded directly; the custom geometry succeeded after supplying path commands in the workflow example. This does not verify every shape-specific adjustment parameter.

### enum.ChartType (25)

Supported chart types for worksheet charts.

`line`, `pie`, `bar`, `doughnut`, `scatter`, `bubble`, `radar`, `treemap`, `sunburst`, `map`, `waterfall`, `line3D`, `pie3D`, `area3D`, `bar3D`, `funnel`, `histogram`, `boxWhisker`, `stock`, `surface3D`, `ofPie`, `surface`, `pareto`, `combo`, `area`.

Map charts currently render country-level data only. Use country names or ISO-3166-1 alpha-2/alpha-3 codes (e.g. USA, CAN, MEX).
### enum.ShapeGeometry (190)

Supported shape geometry tokens for worksheet shapes.

`line`, `lineInv`, `triangle`, `rtTriangle`, `rect`, `diamond`, `parallelogram`, `trapezoid`, `nonIsoscelesTrapezoid`, `pentagon`, `hexagon`, `heptagon`, `octagon`, `decagon`, `dodecagon`, `star4`, `star5`, `star6`, `star7`, `star8`, `star10`, `star12`, `star16`, `star24`, `star32`, `roundRect`, `round1Rect`, `round2SameRect`, `round2DiagRect`, `snipRoundRect`, `snip1Rect`, `snip2SameRect`, `snip2DiagRect`, `plaque`, `ellipse`, `teardrop`, `homePlate`, `chevron`, `pieWedge`, `pie`, `blockArc`, `donut`, `noSmoking`, `rightArrow`, `leftArrow`, `upArrow`, `downArrow`, `stripedRightArrow`, `notchedRightArrow`, `bentUpArrow`, `leftRightArrow`, `upDownArrow`, `leftUpArrow`, `leftRightUpArrow`, `quadArrow`, `leftArrowCallout`, `rightArrowCallout`, `upArrowCallout`, `downArrowCallout`, `leftRightArrowCallout`, `upDownArrowCallout`, `quadArrowCallout`, `bentArrow`, `uturnArrow`, `circularArrow`, `leftCircularArrow`, `leftRightCircularArrow`, `curvedRightArrow`, `curvedLeftArrow`, `curvedUpArrow`, `curvedDownArrow`, `swooshArrow`, `cube`, `can`, `lightningBolt`, `heart`, `sun`, `moon`, `smileyFace`, `irregularSeal1`, `irregularSeal2`, `foldedCorner`, `bevel`, `frame`, `halfFrame`, `corner`, `diagStripe`, `chord`, `arc`, `leftBracket`, `rightBracket`, `leftBrace`, `rightBrace`, `bracketPair`, `bracePair`, `straightConnector1`, `bentConnector2`, `bentConnector3`, `bentConnector4`, `bentConnector5`, `curvedConnector2`, `curvedConnector3`, `curvedConnector4`, `curvedConnector5`, `callout1`, `callout2`, `callout3`, `accentCallout1`, `accentCallout2`, `accentCallout3`, `borderCallout1`, `borderCallout2`, `borderCallout3`, `accentBorderCallout1`, `accentBorderCallout2`, `accentBorderCallout3`, `wedgeRectCallout`, `wedgeRoundRectCallout`, `wedgeEllipseCallout`, `cloudCallout`, `cloud`, `ribbon`, `ribbon2`, `ellipseRibbon`, `ellipseRibbon2`, `leftRightRibbon`, `verticalScroll`, `horizontalScroll`, `wave`, `doubleWave`, `plus`, `flowChartProcess`, `flowChartDecision`, `flowChartInputOutput`, `flowChartPredefinedProcess`, `flowChartInternalStorage`, `flowChartDocument`, `flowChartMultidocument`, `flowChartTerminator`, `flowChartPreparation`, `flowChartManualInput`, `flowChartManualOperation`, `flowChartConnector`, `flowChartPunchedCard`, `flowChartPunchedTape`, `flowChartSummingJunction`, `flowChartOr`, `flowChartCollate`, `flowChartSort`, `flowChartExtract`, `flowChartMerge`, `flowChartOfflineStorage`, `flowChartOnlineStorage`, `flowChartMagneticTape`, `flowChartMagneticDisk`, `flowChartDatabase`, `flowChartMagneticDrum`, `flowChartDisplay`, `flowChartDelay`, `flowChartAlternateProcess`, `flowChartOffpageConnector`, `actionButtonBlank`, `actionButtonHome`, `actionButtonHelp`, `actionButtonInformation`, `actionButtonForwardNext`, `actionButtonBackPrevious`, `actionButtonEnd`, `actionButtonBeginning`, `actionButtonReturn`, `actionButtonDocument`, `actionButtonSound`, `actionButtonMovie`, `gear6`, `gear9`, `funnel`, `mathPlus`, `mathMinus`, `mathMultiply`, `mathDivide`, `mathEqual`, `mathNotEqual`, `cornerTabs`, `squareTabs`, `plaqueTabs`, `chartX`, `chartStar`, `chartPlus`, `custom`, `textbox`.


### enum.LineStyle (5)

Supported line styles for shape and chart strokes.

`solid`, `dashed`, `dotted`, `dash-dot`, `dash-dot-dot`.


<details>
<summary>Additional serialized enum fields and values</summary>

#### Workbook.apply enums (v0)

Shared enum value lists referenced by `Workbook.apply` ops.

##### `props.axis.maxMode`

- `custom`
- `group`
- `individual`
- `unspecified`

##### `props.axis.minMode`

- `custom`
- `group`
- `individual`
- `unspecified`

##### `props.borders.preset`

- `all`
- `doubleBottom`
- `inside`
- `none`
- `outside`

##### `props.chartType`

- `area`
- `area3D`
- `bar`
- `bar3D`
- `boxWhisker`
- `bubble`
- `combo`
- `doughnut`
- `funnel`
- `histogram`
- `line`
- `line3D`
- `map`
- `ofPie`
- `pareto`
- `pie`
- `pie3D`
- `radar`
- `scatter`
- `stock`
- `sunburst`
- `surface`
- `surface3D`
- `treemap`
- `waterfall`

##### `props.displayBlanksAs`

- `gap`
- `span`
- `zero`

##### `props.displayEmptyCellsAs`

- `gap`
- `span`
- `unspecified`
- `zero`

##### `props.errorAlert.style`

- `information`
- `stop`
- `warning`

##### `props.fill.color.value`

- `accent1`
- `accent2`
- `accent3`
- `accent4`
- `accent5`
- `accent6`
- `bg1`
- `bg2`
- `dk1`
- `dk2`
- `folHlink`
- `hlink`
- `lt1`
- `lt2`
- `tx1`
- `tx2`

##### `props.fill.gradientKind`

- `linear`
- `path`

##### `props.fill.pattern.type`

- `cross`
- `darkDown`
- `darkGray`
- `darkGrid`
- `darkHorizontal`
- `darkTrellis`
- `darkUp`
- `darkVertical`
- `dashedDownwardDiagonal`
- `dashedHorizontal`
- `dashedUpwardDiagonal`
- `dashedVertical`
- `diagonalBrick`
- `diagonalCross`
- `divot`
- `dotGrid`
- `dottedDiamond`
- `downwardDiagonal`
- `gray0625`
- `gray125`
- `horizontal`
- `horizontalBrick`
- `largeCheck`
- `largeConfetti`
- `largeGrid`
- `lightDown`
- `lightGray`
- `lightGrid`
- `lightHorizontal`
- `lightTrellis`
- `lightUp`
- `lightVertical`
- `mediumGray`
- `narrowHorizontal`
- `narrowVertical`
- `none`
- `openDiamond`
- `percent10`
- `percent20`
- `percent25`
- `percent30`
- `percent40`
- `percent5`
- `percent50`
- `percent60`
- `percent70`
- `percent75`
- `percent80`
- `percent90`
- `plaid`
- `shingle`
- `smallCheck`
- `smallConfetti`
- `smallGrid`
- `solid`
- `solidDiamond`
- `sphere`
- `trellis`
- `upwardDiagonal`
- `vertical`
- `wave`
- `weave`
- `wideDownwardDiagonal`
- `wideUpwardDiagonal`
- `zigZag`

##### `props.geometry`

- `accentBorderCallout1`
- `accentBorderCallout2`
- `accentBorderCallout3`
- `accentCallout1`
- `accentCallout2`
- `accentCallout3`
- `actionButtonBackPrevious`
- `actionButtonBeginning`
- `actionButtonBlank`
- `actionButtonDocument`
- `actionButtonEnd`
- `actionButtonForwardNext`
- `actionButtonHelp`
- `actionButtonHome`
- `actionButtonInformation`
- `actionButtonMovie`
- `actionButtonReturn`
- `actionButtonSound`
- `arc`
- `bentArrow`
- `bentConnector2`
- `bentConnector3`
- `bentConnector4`
- `bentConnector5`
- `bentUpArrow`
- `bevel`
- `blockArc`
- `borderCallout1`
- `borderCallout2`
- `borderCallout3`
- `bracePair`
- `bracketPair`
- `callout1`
- `callout2`
- `callout3`
- `can`
- `chartPlus`
- `chartStar`
- `chartX`
- `chevron`
- `chord`
- `circularArrow`
- `cloud`
- `cloudCallout`
- `corner`
- `cornerTabs`
- `cube`
- `curvedConnector2`
- `curvedConnector3`
- `curvedConnector4`
- `curvedConnector5`
- `curvedDownArrow`
- `curvedLeftArrow`
- `curvedRightArrow`
- `curvedUpArrow`
- `decagon`
- `diagStripe`
- `diamond`
- `dodecagon`
- `donut`
- `doubleWave`
- `downArrow`
- `downArrowCallout`
- `ellipse`
- `ellipseRibbon`
- `ellipseRibbon2`
- `flowChartAlternateProcess`
- `flowChartCollate`
- `flowChartConnector`
- `flowChartDecision`
- `flowChartDelay`
- `flowChartDisplay`
- `flowChartDocument`
- `flowChartExtract`
- `flowChartInputOutput`
- `flowChartInternalStorage`
- `flowChartMagneticDisk`
- `flowChartMagneticDrum`
- `flowChartMagneticTape`
- `flowChartManualInput`
- `flowChartManualOperation`
- `flowChartMerge`
- `flowChartMultidocument`
- `flowChartOfflineStorage`
- `flowChartOffpageConnector`
- `flowChartOnlineStorage`
- `flowChartOr`
- `flowChartPredefinedProcess`
- `flowChartPreparation`
- `flowChartProcess`
- `flowChartPunchedCard`
- `flowChartPunchedTape`
- `flowChartSort`
- `flowChartSummingJunction`
- `flowChartTerminator`
- `foldedCorner`
- `frame`
- `funnel`
- `gear6`
- `gear9`
- `halfFrame`
- `heart`
- `heptagon`
- `hexagon`
- `homePlate`
- `horizontalScroll`
- `irregularSeal1`
- `irregularSeal2`
- `leftArrow`
- `leftArrowCallout`
- `leftBrace`
- `leftBracket`
- `leftCircularArrow`
- `leftRightArrow`
- `leftRightArrowCallout`
- `leftRightCircularArrow`
- `leftRightRibbon`
- `leftRightUpArrow`
- `leftUpArrow`
- `lightningBolt`
- `line`
- `lineInv`
- `mathDivide`
- `mathEqual`
- `mathMinus`
- `mathMultiply`
- `mathNotEqual`
- `mathPlus`
- `moon`
- `noSmoking`
- `nonIsoscelesTrapezoid`
- `notchedRightArrow`
- `octagon`
- `parallelogram`
- `pentagon`
- `pie`
- `pieWedge`
- `plaque`
- `plaqueTabs`
- `plus`
- `quadArrow`
- `quadArrowCallout`
- `rect`
- `ribbon`
- `ribbon2`
- `rightArrow`
- `rightArrowCallout`
- `rightBrace`
- `rightBracket`
- `round1Rect`
- `round2DiagRect`
- `round2SameRect`
- `roundRect`
- `rtTriangle`
- `smileyFace`
- `snip1Rect`
- `snip2DiagRect`
- `snip2SameRect`
- `snipRoundRect`
- `squareTabs`
- `star10`
- `star12`
- `star16`
- `star24`
- `star32`
- `star4`
- `star5`
- `star6`
- `star7`
- `star8`
- `straightConnector1`
- `stripedRightArrow`
- `sun`
- `swooshArrow`
- `teardrop`
- `textbox`
- `trapezoid`
- `triangle`
- `upArrow`
- `upArrowCallout`
- `upDownArrow`
- `upDownArrowCallout`
- `uturnArrow`
- `verticalScroll`
- `wave`
- `wedgeEllipseCallout`
- `wedgeRectCallout`
- `wedgeRoundRectCallout`

##### `props.horizontalAlignment`

- `Center`
- `CenterAcrossSelection`
- `Distributed`
- `Fill`
- `General`
- `Justify`
- `Left`
- `Right`

##### `props.legend.position`

- `bottom`
- `left`
- `right`
- `top`
- `topRight`

##### `props.line.style`

- `dash-dot`
- `dash-dot-dot`
- `dashed`
- `dotted`
- `solid`

##### `props.rule.format.fill.color.value`

- `accent1`
- `accent2`
- `accent3`
- `accent4`
- `accent5`
- `accent6`
- `bg1`
- `bg2`
- `dk1`
- `dk2`
- `folHlink`
- `hlink`
- `lt1`
- `lt2`
- `tx1`
- `tx2`

##### `props.rule.format.fill.gradientKind`

- `linear`
- `path`

##### `props.rule.format.fill.pattern.type`

- `cross`
- `darkDown`
- `darkGray`
- `darkGrid`
- `darkHorizontal`
- `darkTrellis`
- `darkUp`
- `darkVertical`
- `dashedDownwardDiagonal`
- `dashedHorizontal`
- `dashedUpwardDiagonal`
- `dashedVertical`
- `diagonalBrick`
- `diagonalCross`
- `divot`
- `dotGrid`
- `dottedDiamond`
- `downwardDiagonal`
- `gray0625`
- `gray125`
- `horizontal`
- `horizontalBrick`
- `largeCheck`
- `largeConfetti`
- `largeGrid`
- `lightDown`
- `lightGray`
- `lightGrid`
- `lightHorizontal`
- `lightTrellis`
- `lightUp`
- `lightVertical`
- `mediumGray`
- `narrowHorizontal`
- `narrowVertical`
- `none`
- `openDiamond`
- `percent10`
- `percent20`
- `percent25`
- `percent30`
- `percent40`
- `percent5`
- `percent50`
- `percent60`
- `percent70`
- `percent75`
- `percent80`
- `percent90`
- `plaid`
- `shingle`
- `smallCheck`
- `smallConfetti`
- `smallGrid`
- `solid`
- `solidDiamond`
- `sphere`
- `trellis`
- `upwardDiagonal`
- `vertical`
- `wave`
- `weave`
- `wideDownwardDiagonal`
- `wideUpwardDiagonal`
- `zigZag`

##### `props.rule.operator`

- `between`
- `equal`
- `greaterThan`
- `greaterThanOrEqual`
- `lessThan`
- `lessThanOrEqual`
- `notBetween`
- `notEqual`

##### `props.rule.thresholds[].type`

- `max`
- `min`
- `num`
- `percent`
- `percentile`

##### `props.rule.type`

- `custom`
- `date`
- `decimal`
- `list`
- `none`
- `textLength`
- `time`
- `whole`

##### `props.series[].marker.symbol`

- `circle`
- `diamond`
- `dot`
- `none`
- `plus`
- `square`
- `star`
- `triangle`
- `x`

##### `props.series[].stroke.fill.color.value`

- `accent1`
- `accent2`
- `accent3`
- `accent4`
- `accent5`
- `accent6`
- `bg1`
- `bg2`
- `dk1`
- `dk2`
- `folHlink`
- `hlink`
- `lt1`
- `lt2`
- `tx1`
- `tx2`

##### `props.series[].stroke.fill.gradientKind`

- `linear`
- `path`

##### `props.series[].stroke.fill.pattern.type`

- `cross`
- `darkDown`
- `darkGray`
- `darkGrid`
- `darkHorizontal`
- `darkTrellis`
- `darkUp`
- `darkVertical`
- `dashedDownwardDiagonal`
- `dashedHorizontal`
- `dashedUpwardDiagonal`
- `dashedVertical`
- `diagonalBrick`
- `diagonalCross`
- `divot`
- `dotGrid`
- `dottedDiamond`
- `downwardDiagonal`
- `gray0625`
- `gray125`
- `horizontal`
- `horizontalBrick`
- `largeCheck`
- `largeConfetti`
- `largeGrid`
- `lightDown`
- `lightGray`
- `lightGrid`
- `lightHorizontal`
- `lightTrellis`
- `lightUp`
- `lightVertical`
- `mediumGray`
- `narrowHorizontal`
- `narrowVertical`
- `none`
- `openDiamond`
- `percent10`
- `percent20`
- `percent25`
- `percent30`
- `percent40`
- `percent5`
- `percent50`
- `percent60`
- `percent70`
- `percent75`
- `percent80`
- `percent90`
- `plaid`
- `shingle`
- `smallCheck`
- `smallConfetti`
- `smallGrid`
- `solid`
- `solidDiamond`
- `sphere`
- `trellis`
- `upwardDiagonal`
- `vertical`
- `wave`
- `weave`
- `wideDownwardDiagonal`
- `wideUpwardDiagonal`
- `zigZag`

##### `props.series[].stroke.style`

- `dash-dot`
- `dash-dot-dot`
- `dashed`
- `dotted`
- `solid`

##### `props.seriesColor.value`

- `accent1`
- `accent2`
- `accent3`
- `accent4`
- `accent5`
- `accent6`
- `bg1`
- `bg2`
- `dk1`
- `dk2`
- `folHlink`
- `hlink`
- `lt1`
- `lt2`
- `tx1`
- `tx2`

##### `props.type`

- `column`
- `line`
- `stacked`

##### `props.verticalAlignment`

- `bottom`
- `middle`
- `top`

</details>

## All 494 formulas: arguments, input and observed output

Shared fixture: Sheet1 and Data A1:E20 contain rows [i, 2i, 3i, 4i, 5i] for i=1…20. Each supplied formula runs at Sheet1!Z100. Results below are observed top-left values, not independently certified function specifications. The JSON receipt also captures a 4×4 projection. #NUM!/#VALUE! can reflect missing example prerequisites. “Value returned” is not a correctness verdict.

| Category | Count |
| --- | --- |
| compatibility | 42 |
| cube | 7 |
| database | 12 |
| date-time | 24 |
| engineering | 54 |
| financial | 55 |
| information | 19 |
| logical | 19 |
| lookup-reference | 41 |
| math-trig | 77 |
| statistical | 106 |
| text | 38 |

### Formula category: compatibility

<details>
<summary>STDEV — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Estimates standard deviation based on a sample (ignores logical values and text in the sample)

- number1 (number|range, required): Are 1 to 255 numbers corresponding to a sample of a population and can be numbers or references that contain numbers
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=STDEV(10, C1:C5, 20, C1:C5)`

Observed output: `5.134553180524705`

Classification: value returned. Elapsed: 16 ms.

</details>

<details>
<summary>VAR — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Estimates variance based on a sample (ignores logical values and text in the sample)

- number1 (number|range, required): Are 1 to 255 numeric arguments corresponding to a sample of a population
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=VAR(10, C1:C5, 20, C1:C5)`

Observed output: `26.363636363636363`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>STDEVP — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Calculates standard deviation based on the entire population given as arguments (ignores logical values and text)

- number1 (number|range, required): Are 1 to 255 numbers corresponding to a population and can be numbers or references that contain numbers
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=STDEVP(10, C1:C5, 20, C1:C5)`

Observed output: `4.915960401250875`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>VARP — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Calculates variance based on the entire population (ignores logical values and text in the population)

- number1 (number|range, required): Are 1 to 255 numeric arguments corresponding to a population
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=VARP(10, C1:C5, 20, C1:C5)`

Observed output: `24.166666666666668`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>RANK — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the rank of a number in a list of numbers: its size relative to other values in the list

- numberParam (number|range, required): Is the number for which you want to find the rank
- ref (value|range, required): Is an array of, or a reference to, a list of numbers. Nonnumeric values are ignored
- order (value|range, optional): Is a number: rank in the list sorted descending = 0 or omitted; rank in the list sorted ascending = any nonzero value

Input: `=RANK(10, A1, A1)`

Observed output: `0`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>BETADIST — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the cumulative beta probability density function

- x (value|range, required): Is the value between A and B at which to evaluate the function
- alpha (value|range, required): Is a parameter to the distribution and must be greater than 0
- beta (value|range, required): Is a parameter to the distribution and must be greater than 0
- a (value|range, optional): Is an optional lower bound to the interval of x. If omitted, A = 0
- b (value|range, optional): Is an optional upper bound to the interval of x. If omitted, B = 1

Input: `=BETADIST(A1, A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>BETAINV — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the inverse of the cumulative beta probability density function (BETADIST)

- probability (value|range, required): Is a probability associated with the beta distribution
- alpha (value|range, required): Is a parameter to the distribution and must be greater than 0
- beta (value|range, required): Is a parameter to the distribution and must be greater than 0
- a (value|range, optional): Is an optional lower bound to the interval of x. If omitted, A = 0
- b (value|range, optional): Is an optional upper bound to the interval of x. If omitted, B = 1

Input: `=BETAINV(A1, A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>BINOMDIST — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the individual term binomial distribution probability

- numberS (number|range, required): Is the number of successes in trials
- trials (value|range, required): Is the number of independent trials
- probabilityS (value|range, required): Is the probability of success on each trial
- cumulative (value|range, required): Is a logical value: for the cumulative distribution function, use TRUE; for the probability mass function, use FALSE

Input: `=BINOMDIST(10, A1, A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>CHIDIST — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the right-tailed probability of the chi-squared distribution

- x (value|range, required): Is the value at which you want to evaluate the distribution, a nonnegative number
- degFreedom (value|range, required): Is the number of degrees of freedom, a number between 1 and 10^10, excluding 10^10

Input: `=CHIDIST(A1, A1)`

Observed output: `0.24197072451913293`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>CHIINV — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the inverse of the right-tailed probability of the chi-squared distribution

- probability (value|range, required): Is a probability associated with the chi-squared distribution, a value between 0 and 1 inclusive
- degFreedom (value|range, required): Is the number of degrees of freedom, a number between 1 and 10^10, excluding 10^10

Input: `=CHIINV(A1, A1)`

Observed output: `200`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>CONFIDENCE — unimplemented text</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the confidence interval for a population mean, using a normal distribution

- alpha (value|range, required): Is the significance level used to compute the confidence level, a number greater than 0 and less than 1
- standardDev (value|range, required): Is the population standard deviation for the data range and is assumed to be known. Standard_dev must be greater than 0
- size (value|range, required): Is the sample size

Input: `=CONFIDENCE(A1, A1, A1)`

Observed output: `"CONFIDENCE is not implemented. alpha=1, standardDev=1, size=1"`

Classification: unimplemented text. Elapsed: 10 ms.

</details>

<details>
<summary>CRITBINOM — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value

- trials (value|range, required): Is the number of Bernoulli trials
- probabilityS (value|range, required): Is the probability of success on each trial, a number between 0 and 1 inclusive
- alpha (value|range, required): Is the criterion value, a number between 0 and 1 inclusive

Input: `=CRITBINOM(A1, A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>EXPONDIST — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the exponential distribution

- x (value|range, required): Is the value of the function, a nonnegative number
- lambda (value|range, required): Is the parameter value, a positive number
- cumulative (value|range, required): Is a logical value for the function to return: the cumulative distribution function = TRUE; the probability density function = FALSE

Input: `=EXPONDIST(A1, A1, A1)`

Observed output: `0.6321205588285577`

Classification: value returned. Elapsed: 6 ms.

</details>

<details>
<summary>FDIST — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the (right-tailed) F probability distribution (degree of diversity) for two data sets

- x (value|range, required): Is the value at which to evaluate the function, a nonnegative number
- degFreedom1 (value|range, required): Is the numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10
- degFreedom2 (value|range, required): Is the denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10

Input: `=FDIST(A1, A1, A1)`

Observed output: `0.1591549421985171`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>FINV — Excel error</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the inverse of the (right-tailed) F probability distribution: if p = FDIST(x,...), then FINV(p,...) = x

- probability (value|range, required): Is a probability associated with the F cumulative distribution, a number between 0 and 1 inclusive
- degFreedom1 (value|range, required): Is the numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10
- degFreedom2 (value|range, required): Is the denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10

Input: `=FINV(A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 9 ms.

</details>

<details>
<summary>FLOOR — value returned</summary>

Rounds a number down to the nearest multiple of significance

- numberParam (number|range, required): Is the numeric value you want to round
- significance (value|range, required): Is the multiple to which you want to round. Number and Significance must either both be positive or both be negative

Input: `=FLOOR(10, A1)`

Observed output: `10`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>GAMMADIST — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the gamma distribution

- x (value|range, required): Is the value at which you want to evaluate the distribution, a nonnegative number
- alpha (value|range, required): Is a parameter to the distribution, a positive number
- beta (value|range, required): Is a parameter to the distribution, a positive number. If beta = 1, GAMMADIST returns the standard gamma distribution
- cumulative (value|range, required): Is a logical value: return the cumulative distribution function = TRUE; return the probability mass function = FALSE or omitted

Input: `=GAMMADIST(A1, A1, A1, A1)`

Observed output: `0.6321205588285578`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>GAMMAINV — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the inverse of the gamma cumulative distribution: if p = GAMMADIST(x,...), then GAMMAINV(p,...) = x

- probability (value|range, required): Is the probability associated with the gamma distribution, a number between 0 and 1, inclusive
- alpha (value|range, required): Is a parameter to the distribution, a positive number
- beta (value|range, required): Is a parameter to the distribution, a positive number. If beta = 1, GAMMAINV returns the inverse of the standard gamma distribution

Input: `=GAMMAINV(A1, A1, A1)`

Observed output: `101`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>CEILING — value returned</summary>

Rounds a number up, to the nearest multiple of significance

- numberParam (number|range, required): Is the value you want to round
- significance (value|range, required): Is the multiple to which you want to round

Input: `=CEILING(10, A1)`

Observed output: `10`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>HYPGEOMDIST — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the hypergeometric distribution

- sampleS (value|range, required): Is the number of successes in the sample
- numberSample (number|range, required): Is the size of the sample
- populationS (value|range, required): Is the number of successes in the population
- numberPop (number|range, required): Is the population size

Input: `=HYPGEOMDIST(A1, 10, A1, 10)`

Observed output: `1`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>LOGNORMDIST — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the cumulative lognormal distribution of x, where ln(x) is normally distributed with parameters Mean and Standard_dev

- x (value|range, required): Is the value at which to evaluate the function, a positive number
- mean (value|range, required): Is the mean of ln(x)
- standardDev (value|range, required): Is the standard deviation of ln(x), a positive number

Input: `=LOGNORMDIST(A1, A1, A1)`

Observed output: `0.24197072451914337`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>LOGINV — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the inverse of the lognormal cumulative distribution function of x, where ln(x) is normally distributed with parameters Mean and Standard_dev

- probability (value|range, required): Is a probability associated with the lognormal distribution, a number between 0 and 1, inclusive
- mean (value|range, required): Is the mean of ln(x)
- standardDev (value|range, required): Is the standard deviation of ln(x), a positive number

Input: `=LOGINV(A1, A1, A1)`

Observed output: `7.125397860548684e+61`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>NEGBINOMDIST — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the negative binomial distribution, the probability that there will be Number_f failures before the Number_s-th success, with Probability_s probability of a success

- numberF (number|range, required): Is the number of failures
- numberS (number|range, required): Is the threshold number of successes
- probabilityS (value|range, required): Is the probability of a success; a number between 0 and 1

Input: `=NEGBINOMDIST(10, 10, A1)`

Observed output: `0`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>NORMDIST — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the normal cumulative distribution for the specified mean and standard deviation

- x (value|range, required): Is the value for which you want the distribution
- mean (value|range, required): Is the arithmetic mean of the distribution
- standardDev (value|range, required): Is the standard deviation of the distribution, a positive number
- cumulative (value|range, required): Is a logical value: for the cumulative distribution function, use TRUE; for the probability density function, use FALSE

Input: `=NORMDIST(A1, A1, A1, A1)`

Observed output: `0.5`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>NORMSDIST — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the standard normal cumulative distribution (has a mean of zero and a standard deviation of one)

- z (value|range, required): Is the value for which you want the distribution

Input: `=NORMSDIST(A1)`

Observed output: `0.8413447460685429`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>NORMINV — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation

- probability (value|range, required): Is a probability corresponding to the normal distribution, a number between 0 and 1 inclusive
- mean (value|range, required): Is the arithmetic mean of the distribution
- standardDev (value|range, required): Is the standard deviation of the distribution, a positive number

Input: `=NORMINV(A1, A1, A1)`

Observed output: `142.4213562373095`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>NORMSINV — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the inverse of the standard normal cumulative distribution (has a mean of zero and a standard deviation of one)

- probability (value|range, required): Is a probability corresponding to the normal distribution, a number between 0 and 1 inclusive

Input: `=NORMSINV(A1)`

Observed output: `141.4213562373095`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>POISSON — unimplemented text</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the Poisson distribution

- x (value|range, required): Is the number of events
- mean (value|range, required): Is the expected numeric value, a positive number
- cumulative (value|range, required): Is a logical value: for the cumulative Poisson probability, use TRUE; for the Poisson probability mass function, use FALSE

Input: `=POISSON(A1, A1, A1)`

Observed output: `"POISSON is not implemented. x=1, mean=1, cumulative=1"`

Classification: unimplemented text. Elapsed: 9 ms.

</details>

<details>
<summary>TDIST — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the Student's t-distribution

- x (value|range, required): Is the numeric value at which to evaluate the distribution
- degFreedom (value|range, required): Is an integer indicating the number of degrees of freedom that characterize the distribution
- tails (value|range, required): Specifies the number of distribution tails to return: one-tailed distribution = 1; two-tailed distribution = 2

Input: `=TDIST(A1, A1, A1)`

Observed output: `0.24999999852683763`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>WEIBULL — unimplemented text</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the Weibull distribution

- x (value|range, required): Is the value at which to evaluate the function, a nonnegative number
- alpha (value|range, required): Is a parameter to the distribution, a positive number
- beta (value|range, required): Is a parameter to the distribution, a positive number
- cumulative (value|range, required): Is a logical value: for the cumulative distribution function, use TRUE; for the probability mass function, use FALSE

Input: `=WEIBULL(A1, A1, A1, A1)`

Observed output: `"WEIBULL is not implemented. x=1, alpha=1, beta=1, cumulative=1"`

Classification: unimplemented text. Elapsed: 10 ms.

</details>

<details>
<summary>CHITEST — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the test for independence: the value from the chi-squared distribution for the statistic and the appropriate degrees of freedom

- actualRange (range, required): Is the range of data that contains observations to test against expected values
- expectedRange (range, required): Is the range of data that contains the ratio of the product of row totals and column totals to the grand total

Input: `=CHITEST(A1:A5, A1:A5)`

Observed output: `1`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>COVAR — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns covariance, the average of the products of deviations for each data point pair in two data sets

- array1 (array, required): Is the first cell range of integers and must be numbers, arrays, or references that contain numbers
- array2 (array, required): Is the second cell range of integers and must be numbers, arrays, or references that contain numbers

Input: `=COVAR({1,2;3,4}, {1,2;3,4})`

Observed output: `1.25`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>FORECAST — Excel error</summary>

This function is available for compatibility with Excel 2013 and earlier. Calculates, or predicts, a future value along a linear trend by using existing values

- x (value|range, required): Is the data point for which you want to predict a value and must be a numeric value
- knownYs (value|range, required): Is the dependent array or range of numeric data
- knownXs (value|range, required): Is the independent array or range of numeric data. The variance of Known_x's must not be zero

Input: `=FORECAST(A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 11 ms.

</details>

<details>
<summary>FTEST — Excel error</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the result of an F-test, the two-tailed probability that the variances in Array1 and Array2 are not significantly different

- array1 (array, required): Is the first array or range of data and can be numbers or names, arrays, or references that contain numbers (blanks are ignored)
- array2 (array, required): Is the second array or range of data and can be numbers or names, arrays, or references that contain numbers (blanks are ignored)

Input: `=FTEST({1,2;3,4}, {1,2;3,4})`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 12 ms.

</details>

<details>
<summary>TTEST — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the probability associated with a Student's t-Test

- array1 (array, required): Is the first data set
- array2 (array, required): Is the second data set
- tails (value|range, required): Specifies the number of distribution tails to return: one-tailed distribution = 1; two-tailed distribution = 2
- typeParam (value|range, required): Is the kind of t-test: paired = 1, two-sample equal variance (homoscedastic) = 2, two-sample unequal variance = 3

Input: `=TTEST({1,2;3,4}, {1,2;3,4}, A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>ZTEST — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the one-tailed P-value of a z-test

- array (array, required): Is the array or range of data against which to test X
- x (value|range, required): Is the value to test
- sigma (value|range, optional): Is the population (known) standard deviation. If omitted, the sample standard deviation is used

Input: `=ZTEST({1,2;3,4}, A1, A1)`

Observed output: `0.0013498980316301035`

Classification: value returned. Elapsed: 16 ms.

</details>

<details>
<summary>QUARTILE — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the quartile of a data set

- array (array, required): Is the array or cell range of numeric values for which you want the quartile value
- quart (value|range, required): Is a number: minimum value = 0; 1st quartile = 1; median value = 2; 3rd quartile = 3; maximum value = 4

Input: `=QUARTILE({1,2;3,4}, A1)`

Observed output: `1.75`

Classification: value returned. Elapsed: 15 ms.

</details>

<details>
<summary>PERCENTILE — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the k-th percentile of values in a range

- array (array, required): Is the array or range of data that defines relative standing
- k (value|range, required): Is the percentile value that is between 0 through 1, inclusive

Input: `=PERCENTILE({1,2;3,4}, A1)`

Observed output: `4`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>PERCENTRANK — value returned</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the rank of a value in a data set as a percentage of the data set

- array (array, required): Is the array or range of data with numeric values that defines relative standing
- x (value|range, required): Is the value for which you want to know the rank
- significance (value|range, optional): Is an optional value that identifies the number of significant digits for the returned percentage, three digits if omitted (0.xxx%)

Input: `=PERCENTRANK({1,2;3,4}, A1, A1)`

Observed output: `0`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>MODE — unimplemented text</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the most frequently occurring, or repetitive, value in an array or range of data

- number1 (number|range, required): Are 1 to 255 numbers, or names, arrays, or references that contain numbers for which you want the mode
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=MODE(10, C1:C5, 20, C1:C5)`

Observed output: `"MODE is not implemented. number1=10, number2=3,6,9,12,15, rest.length=2"`

Classification: unimplemented text. Elapsed: 12 ms.

</details>

<details>
<summary>TINV — Excel error</summary>

This function is available for compatibility with Excel 2007 and earlier. Returns the two-tailed inverse of the Student's t-distribution

- probability (value|range, required): Is the probability associated with the two-tailed Student's t-distribution, a number between 0 and 1 inclusive
- degFreedom (value|range, required): Is a positive integer indicating the number of degrees of freedom to characterize the distribution

Input: `=TINV(A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 8 ms.

</details>

<details>
<summary>CONCATENATE — value returned</summary>

Joins several text strings into one text string

- text1 (string, required): Are 1 to 255 text strings to be joined into a single text string and can be text strings, numbers, or single-cell references
- text2 (string, optional): Text string.
- rest (value|range, required): Value, reference, or range.

Input: `=CONCATENATE("text1", C1:C5, "text2", C1:C5)`

Observed output: `"text13691215text23691215"`

Classification: value returned. Elapsed: 10 ms.

</details>


### Formula category: cube

<details>
<summary>CUBEVALUE — unimplemented text</summary>

Returns an aggregated value from the cube.

- connection (value|range, required): Is the name of a connection to an OLAP cube
- memberExpression1 (value|range, optional): Is a slicer that determines the portion of the OLAP cube for which the aggregated value is to be retrieved
- rest (value|range, required): Value, reference, or range.

Input: `=CUBEVALUE(A1, C1:C5, A1, C1:C5)`

Observed output: `"CUBEVALUE is not implemented. connection=1, memberExpression1=3,6,9,12,15, rest.length=2"`

Classification: unimplemented text. Elapsed: 8 ms.

</details>

<details>
<summary>CUBEMEMBER — unimplemented text</summary>

Returns a member or tuple from the cube.

- connection (value|range, required): Is the name of a connection to an OLAP cube
- memberExpression (value|range, required): Is the expression representing the name of a member or tuple in the OLAP cube
- caption (value|range, optional): Is the caption to be displayed in the cell

Input: `=CUBEMEMBER(A1, A1, A1)`

Observed output: `"CUBEMEMBER is not implemented. connection=1, memberExpression=1, caption=1"`

Classification: unimplemented text. Elapsed: 8 ms.

</details>

<details>
<summary>CUBEMEMBERPROPERTY — unimplemented text</summary>

Returns the value of a member property from the cube.

- connection (value|range, required): Is the name of a connection to an OLAP cube
- memberExpression (value|range, required): Is the expression representing the name of a member in the OLAP cube
- property (value|range, required): Is the property name

Input: `=CUBEMEMBERPROPERTY(A1, A1, A1)`

Observed output: `"CUBEMEMBERPROPERTY is not implemented. connection=1, memberExpression=1, property=1"`

Classification: unimplemented text. Elapsed: 9 ms.

</details>

<details>
<summary>CUBERANKEDMEMBER — unimplemented text</summary>

Returns the nth, or ranked, member in a set.

- connection (value|range, required): Is the name of a connection to an OLAP cube
- setExpression (value|range, required): Is the set from which the element is to be retrieved
- rank (value|range, required): Is the rank of the element to be retrieved
- caption (value|range, optional): Is the caption to be displayed in the cell

Input: `=CUBERANKEDMEMBER(A1, A1, A1)`

Observed output: `"CUBERANKEDMEMBER is not implemented. connection=1, setExpression=1, rank=1, caption=undefined"`

Classification: unimplemented text. Elapsed: 9 ms.

</details>

<details>
<summary>CUBEKPIMEMBER — unimplemented text</summary>

Returns a key performance indicator (KPI) property and displays the KPI name in the cell.

- connection (value|range, required): Is the name of a connection to an OLAP cube
- kpiName (value|range, required): Is the KPI name
- kpiProperty (value|range, required): Is the KPI property
- caption (value|range, optional): Is the caption to be displayed in the cell

Input: `=CUBEKPIMEMBER(A1, A1, A1)`

Observed output: `"CUBEKPIMEMBER is not implemented. connection=1, kpiName=1, kpiProperty=1, caption=undefined"`

Classification: unimplemented text. Elapsed: 10 ms.

</details>

<details>
<summary>CUBESET — unimplemented text</summary>

Defines a calculated set of members or tuples by sending a set expression to the cube on the server, which creates the set, and then returns that set to Microsoft Excel.

- connection (value|range, required): Is the name of a connection to an OLAP cube
- setExpression (value|range, required): Is the expression for the set
- caption (value|range, optional): Is the caption to be displayed in the cell
- sortOrder (value|range, optional): Is the sort order
- sortBy (value|range, optional): Is the sort by

Input: `=CUBESET(A1, A1, A1)`

Observed output: `"CUBESET is not implemented. connection=1, setExpression=1, caption=1, sortOrder=undefined, sortBy=undefined"`

Classification: unimplemented text. Elapsed: 8 ms.

</details>

<details>
<summary>CUBESETCOUNT — unimplemented text</summary>

Returns the number of items in a set.

- setParam (value|range, required): Is the set whose elements are to be counted

Input: `=CUBESETCOUNT(A1)`

Observed output: `"CUBESETCOUNT is not implemented. setParam=1"`

Classification: unimplemented text. Elapsed: 7 ms.

</details>


### Formula category: database

<details>
<summary>DCOUNT — exception-like text</summary>

Counts the cells containing numbers in the field (column) of records in the database that match the conditions you specify

- database (value|range, required): Is the range of cells that makes up the list or database. A database is a list of related data
- field (value|range, required): Is either the label of the column in double quotation marks or a number that represents the column's position in the list
- criteria (criteria, required): Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition

Input: `=DCOUNT(A1, A1, ">0")`

Observed output: `"Cannot read properties of undefined (reading 'length')"`

Classification: exception-like text. Elapsed: 8 ms.

</details>

<details>
<summary>DSUM — value returned</summary>

Adds the numbers in the field (column) of records in the database that match the conditions you specify

- database (value|range, required): Is the range of cells that makes up the list or database. A database is a list of related data
- field (value|range, required): Is either the label of the column in double quotation marks or a number that represents the column's position in the list
- criteria (criteria, required): Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition

Input: `=DSUM(A1, A1, ">0")`

Observed output: `0`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>DAVERAGE — exception-like text</summary>

Returns the average of the values in a database field after filtering rows with a criteria table.

- database (range, required): Structured table (including headers) that holds the records to query.
- field (string|number, required): Column to average, specified by header name ("Sales") or column index.
- criteria (range, required): Criteria range with column labels and the filters to apply to the database.

Input: `=DAVERAGE(A1:C10, "Sales", E1:F2)`

Observed output: `"Cannot read properties of undefined (reading '0')"`

Classification: exception-like text. Elapsed: 9 ms.

</details>

<details>
<summary>DMIN — exception-like text</summary>

Returns the smallest number in the field (column) of records in the database that match the conditions you specify

- database (value|range, required): Is the range of cells that makes up the list or database. A database is a list of related data
- field (value|range, required): Is either the label of the column in double quotation marks or a number that represents the column's position in the list
- criteria (criteria, required): Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition

Input: `=DMIN(A1, A1, ">0")`

Observed output: `"Cannot read properties of undefined (reading 'length')"`

Classification: exception-like text. Elapsed: 9 ms.

</details>

<details>
<summary>DMAX — exception-like text</summary>

Returns the largest number in the field (column) of records in the database that match the conditions you specify

- database (value|range, required): Is the range of cells that makes up the list or database. A database is a list of related data
- field (value|range, required): Is either the label of the column in double quotation marks or a number that represents the column's position in the list
- criteria (criteria, required): Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition

Input: `=DMAX(A1, A1, ">0")`

Observed output: `"Cannot read properties of undefined (reading 'length')"`

Classification: exception-like text. Elapsed: 10 ms.

</details>

<details>
<summary>DSTDEV — exception-like text</summary>

Estimates the standard deviation based on a sample from selected database entries

- database (value|range, required): Is the range of cells that makes up the list or database. A database is a list of related data
- field (value|range, required): Is either the label of the column in double quotation marks or a number that represents the column's position in the list
- criteria (criteria, required): Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition

Input: `=DSTDEV(A1, A1, ">0")`

Observed output: `"Cannot read properties of undefined (reading 'length')"`

Classification: exception-like text. Elapsed: 13 ms.

</details>

<details>
<summary>DVAR — exception-like text</summary>

Estimates variance based on a sample from selected database entries

- database (value|range, required): Is the range of cells that makes up the list or database. A database is a list of related data
- field (value|range, required): Is either the label of the column in double quotation marks or a number that represents the column's position in the list
- criteria (criteria, required): Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition

Input: `=DVAR(A1, A1, ">0")`

Observed output: `"Cannot read properties of undefined (reading 'length')"`

Classification: exception-like text. Elapsed: 13 ms.

</details>

<details>
<summary>DPRODUCT — exception-like text</summary>

Multiplies the values in the field (column) of records in the database that match the conditions you specify

- database (value|range, required): Is the range of cells that makes up the list or database. A database is a list of related data
- field (value|range, required): Is either the label of the column in double quotation marks or a number that represents the column's position in the list
- criteria (criteria, required): Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition

Input: `=DPRODUCT(A1, A1, ">0")`

Observed output: `"Cannot read properties of undefined (reading 'length')"`

Classification: exception-like text. Elapsed: 14 ms.

</details>

<details>
<summary>DSTDEVP — exception-like text</summary>

Calculates the standard deviation based on the entire population of selected database entries

- database (value|range, required): Is the range of cells that makes up the list or database. A database is a list of related data
- field (value|range, required): Is either the label of the column in double quotation marks or a number that represents the column's position in the list
- criteria (criteria, required): Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition

Input: `=DSTDEVP(A1, A1, ">0")`

Observed output: `"Cannot read properties of undefined (reading 'length')"`

Classification: exception-like text. Elapsed: 10 ms.

</details>

<details>
<summary>DVARP — exception-like text</summary>

Calculates variance based on the entire population of selected database entries

- database (value|range, required): Is the range of cells that makes up the list or database. A database is a list of related data
- field (value|range, required): Is either the label of the column in double quotation marks or a number that represents the column's position in the list
- criteria (criteria, required): Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition

Input: `=DVARP(A1, A1, ">0")`

Observed output: `"Cannot read properties of undefined (reading 'length')"`

Classification: exception-like text. Elapsed: 16 ms.

</details>

<details>
<summary>DCOUNTA — exception-like text</summary>

Counts nonblank cells in the field (column) of records in the database that match the conditions you specify

- database (value|range, required): Is the range of cells that makes up the list or database. A database is a list of related data
- field (value|range, required): Is either the label of the column in double quotation marks or a number that represents the column's position in the list
- criteria (criteria, required): Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition

Input: `=DCOUNTA(A1, A1, ">0")`

Observed output: `"Cannot read properties of undefined (reading 'length')"`

Classification: exception-like text. Elapsed: 11 ms.

</details>

<details>
<summary>DGET — exception-like text</summary>

Extracts from a database a single record that matches the conditions you specify

- database (value|range, required): Is the range of cells that makes up the list or database. A database is a list of related data
- field (value|range, required): Is either the label of the column in double quotation marks or a number that represents the column's position in the list
- criteria (criteria, required): Is the range of cells that contains the conditions you specify. The range includes a column label and one cell below the label for a condition

Input: `=DGET(A1, A1, ">0")`

Observed output: `"Cannot read properties of undefined (reading 'length')"`

Classification: exception-like text. Elapsed: 8 ms.

</details>


### Formula category: date-time

<details>
<summary>DATE — value returned</summary>

Returns the number that represents the date in Microsoft Excel date-time code

- year (value|range, required): Is a number from 1900 or 1904 (depending on the workbook's date system) to 9999
- month (value|range, required): Is a number from 1 to 12 representing the month of the year
- day (value|range, required): Is a number from 1 to 31 representing the day of the month

Input: `=DATE(A1, A1, A1)`

Observed output: `367`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>TIME — value returned</summary>

Converts hours, minutes, and seconds given as numbers to an Excel serial number, formatted with a time format

- hour (value|range, required): Is a number from 0 to 23 representing the hour
- minute (value|range, required): Is a number from 0 to 59 representing the minute
- second (value|range, required): Is a number from 0 to 59 representing the second

Input: `=TIME(A1, A1, A1)`

Observed output: `0.04237268518518519`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>DAY — value returned</summary>

Returns the day of the month, a number from 1 to 31.

- serialNumber (number|range, required): Is a number in the date-time code used by Microsoft Excel

Input: `=DAY(10)`

Observed output: `10`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>MONTH — value returned</summary>

Returns the month, a number from 1 (January) to 12 (December).

- serialNumber (number|range, required): Is a number in the date-time code used by Microsoft Excel

Input: `=MONTH(10)`

Observed output: `1`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>YEAR — value returned</summary>

Returns the year of a date, an integer in the range 1900 - 9999.

- serialNumber (number|range, required): Is a number in the date-time code used by Microsoft Excel

Input: `=YEAR(10)`

Observed output: `1900`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>WEEKDAY — value returned</summary>

Returns a number from 1 to 7 identifying the day of the week of a date.

- serialNumber (number|range, required): Is a number that represents a date
- returnType (value|range, optional): Is a number: for Sunday=1 through Saturday=7, use 1; for Monday=1 through Sunday=7, use 2; for Monday=0 through Sunday=6, use 3

Input: `=WEEKDAY(10, A1)`

Observed output: `4`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>HOUR — value returned</summary>

Returns the hour as a number from 0 (12:00 A.M.) to 23 (11:00 P.M.).

- serialNumber (number|range, required): Is a number in the date-time code used by Microsoft Excel, or text in time format, such as 16:48:00 or 4:48:00 PM

Input: `=HOUR(10)`

Observed output: `0`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>MINUTE — value returned</summary>

Returns the minute, a number from 0 to 59.

- serialNumber (number|range, required): Is a number in the date-time code used by Microsoft Excel or text in time format, such as 16:48:00 or 4:48:00 PM

Input: `=MINUTE(10)`

Observed output: `0`

Classification: value returned. Elapsed: 6 ms.

</details>

<details>
<summary>SECOND — value returned</summary>

Returns the second, a number from 0 to 59.

- serialNumber (number|range, required): Is a number in the date-time code used by Microsoft Excel or text in time format, such as 16:48:23 or 4:48:47 PM

Input: `=SECOND(10)`

Observed output: `0`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>NOW — value returned</summary>

Returns the current date and time formatted as a date and time.


Input: `=NOW()`

Observed output: `46280`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>DATEVALUE — Excel error</summary>

Converts a date in the form of text to a number that represents the date in Microsoft Excel date-time code

- dateText (string, required): Is text that represents a date in a Microsoft Excel date format, between 1/1/1900 or 1/1/1904 (depending on the workbook's date system) and 12/31/9999

Input: `=DATEVALUE(DATE(2024, 1, 1))`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 6 ms.

</details>

<details>
<summary>TIMEVALUE — value returned</summary>

Converts a text time to an Excel serial number for a time, a number from 0 (12:00:00 AM) to 0.999988426 (11:59:59 PM). Format the number with a time format after entering the formula

- timeText (string, required): Is a text string that gives a time in any one of the Microsoft Excel time formats (date information in the string is ignored)

Input: `=TIMEVALUE(TIME(9, 0, 0))`

Observed output: `0.375`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>DAYS360 — value returned</summary>

Returns the number of days between two dates based on a 360-day year (twelve 30-day months)

- startDate (date, required): Start_date and end_date are the two dates between which you want to know the number of days
- endDate (date, required): Start_date and end_date are the two dates between which you want to know the number of days
- method (value|range, optional): Is a logical value specifying the calculation method: U.S. (NASD) = FALSE or omitted; European = TRUE.

Input: `=DAYS360(DATE(2024, 1, 1), DATE(2024, 1, 1), A1)`

Observed output: `0`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>TODAY — value returned</summary>

Returns the current date formatted as a date.


Input: `=TODAY()`

Observed output: `46280`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>EDATE — value returned</summary>

Returns the serial number of the date that is the indicated number of months before or after the start date

- startDate (date, required): Is a serial date number that represents the start date
- months (value|range, required): Is the number of months before or after start_date

Input: `=EDATE(DATE(2024, 1, 1), A1)`

Observed output: `45323`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>EOMONTH — value returned</summary>

Returns the serial number of the last day of the month before or after a specified number of months

- startDate (date, required): Is a serial date number that represents the start date
- months (value|range, required): Is the number of months before or after the start_date

Input: `=EOMONTH(DATE(2024, 1, 1), A1)`

Observed output: `45351`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>YEARFRAC — value returned</summary>

Returns the year fraction representing the number of whole days between start_date and end_date

- startDate (date, required): Is a serial date number that represents the start date
- endDate (date, required): Is a serial date number that represents the end date
- basis (value|range, optional): Is the type of day count basis to use

Input: `=YEARFRAC(DATE(2024, 1, 1), DATE(2024, 1, 1), A1)`

Observed output: `0`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>WEEKNUM — value returned</summary>

Returns the week number in the year

- serialNumber (number|range, required): Is the date-time code used by Microsoft Excel for date and time calculation
- returnType (value|range, optional): Is a number (1 or 2) that determines the type of the return value

Input: `=WEEKNUM(10, A1)`

Observed output: `2`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>WORKDAY — value returned</summary>

Returns the serial number of the date before or after a specified number of workdays

- startDate (date, required): Is a serial date number that represents the start date
- days (value|range, required): Is the number of nonweekend and non-holiday days before or after start_date
- holidays (value|range, optional): Is an optional array of one or more serial date numbers to exclude from the working calendar, such as state and federal holidays and floating holidays

Input: `=WORKDAY(DATE(2024, 1, 1), A1, A1)`

Observed output: `45293`

Classification: value returned. Elapsed: 16 ms.

</details>

<details>
<summary>NETWORKDAYS — value returned</summary>

Returns the number of whole workdays between two dates

- startDate (date, required): Is a serial date number that represents the start date
- endDate (date, required): Is a serial date number that represents the end date
- holidays (value|range, optional): Is an optional set of one or more serial date numbers to exclude from the working calendar, such as state and federal holidays and floating holidays

Input: `=NETWORKDAYS(DATE(2024, 1, 1), DATE(2024, 1, 1), A1)`

Observed output: `1`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>NETWORKDAYS.INTL — value returned</summary>

Returns the number of whole workdays between two dates with custom weekend parameters

- startDate (date, required): Is a serial date number that represents the start date
- endDate (date, required): Is a serial date number that represents the end date
- weekend (value|range, optional): Is a number or string specifying when weekends occur
- holidays (value|range, optional): Is an optional set of one or more serial date numbers to exclude from the working calendar, such as state and federal holidays and floating holidays

Input: `=NETWORKDAYS.INTL(DATE(2024, 1, 1), DATE(2024, 1, 1), A1)`

Observed output: `1`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>WORKDAY.INTL — value returned</summary>

Returns the serial number of the date before or after a specified number of workdays with custom weekend parameters

- startDate (date, required): Is a serial date number that represents the start date
- days (value|range, required): Is the number of nonweekend and non-holiday days before or after start_date
- weekend (value|range, optional): Is a number or string specifying when weekends occur
- holidays (value|range, optional): Is an optional array of one or more serial date numbers to exclude from the working calendar, such as state and federal holidays and floating holidays

Input: `=WORKDAY.INTL(DATE(2024, 1, 1), A1, A1)`

Observed output: `45293`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>DAYS — value returned</summary>

Returns the number of days between the two dates.

- endDate (date, required): Start_date and end_date are the two dates between which you want to know the number of days
- startDate (date, required): Start_date and end_date are the two dates between which you want to know the number of days

Input: `=DAYS(DATE(2024, 1, 1), DATE(2024, 1, 1))`

Observed output: `0`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>ISOWEEKNUM — value returned</summary>

Returns the ISO week number in the year for a given date

- date (date, required): Is the date-time code used by Microsoft Excel for date and time calculation

Input: `=ISOWEEKNUM(DATE(2024, 1, 1))`

Observed output: `1`

Classification: value returned. Elapsed: 9 ms.

</details>


### Formula category: engineering

<details>
<summary>HEX2BIN — Excel error</summary>

Converts a Hexadecimal number to binary

- numberParam (number|range, required): Is the hexadecimal number you want to convert
- places (value|range, optional): Is the number of characters to use

Input: `=HEX2BIN(10, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 10 ms.

</details>

<details>
<summary>HEX2DEC — value returned</summary>

Converts a hexadecimal number to decimal

- numberParam (number|range, required): Is the hexadecimal number you want to convert

Input: `=HEX2DEC(10)`

Observed output: `16`

Classification: value returned. Elapsed: 6 ms.

</details>

<details>
<summary>HEX2OCT — Excel error</summary>

Converts a hexadecimal number to octal

- numberParam (number|range, required): Is the hexadecimal number you want to convert
- places (value|range, optional): Is the number of characters to use

Input: `=HEX2OCT(10, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 11 ms.

</details>

<details>
<summary>DEC2BIN — Excel error</summary>

Converts a decimal number to binary

- numberParam (number|range, required): Is the decimal integer you want to convert
- places (value|range, optional): Is the number of characters to use

Input: `=DEC2BIN(10, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 7 ms.

</details>

<details>
<summary>DEC2HEX — value returned</summary>

Converts a decimal number to hexadecimal

- numberParam (number|range, required): Is the decimal integer you want to convert
- places (value|range, optional): Is the number of characters to use

Input: `=DEC2HEX(10, A1)`

Observed output: `"a"`

Classification: value returned. Elapsed: 6 ms.

</details>

<details>
<summary>DEC2OCT — Excel error</summary>

Converts a decimal number to octal

- numberParam (number|range, required): Is the decimal integer you want to convert
- places (value|range, optional): Is the number of characters to use

Input: `=DEC2OCT(10, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 7 ms.

</details>

<details>
<summary>OCT2BIN — Excel error</summary>

Converts an octal number to binary

- numberParam (number|range, required): Is the octal number you want to convert
- places (value|range, optional): Is the number of characters to use

Input: `=OCT2BIN(10, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 10 ms.

</details>

<details>
<summary>OCT2HEX — value returned</summary>

Converts an octal number to hexadecimal

- numberParam (number|range, required): Is the octal number you want to convert
- places (value|range, optional): Is the number of characters to use

Input: `=OCT2HEX(10, A1)`

Observed output: `"8"`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>OCT2DEC — value returned</summary>

Converts an octal number to decimal

- numberParam (number|range, required): Is the octal number you want to convert

Input: `=OCT2DEC(10)`

Observed output: `8`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>BIN2DEC — value returned</summary>

Converts a binary number to decimal

- numberParam (number|range, required): Is the binary number you want to convert

Input: `=BIN2DEC(10)`

Observed output: `2`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>BIN2OCT — value returned</summary>

Converts a binary number to octal

- numberParam (number|range, required): Is the binary number you want to convert
- places (value|range, optional): Is the number of characters to use

Input: `=BIN2OCT(10, A1)`

Observed output: `"2"`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>BIN2HEX — value returned</summary>

Converts a binary number to hexadecimal

- numberParam (number|range, required): Is the binary number you want to convert
- places (value|range, optional): Is the number of characters to use

Input: `=BIN2HEX(10, A1)`

Observed output: `"2"`

Classification: value returned. Elapsed: 16 ms.

</details>

<details>
<summary>IMSUB — exception-like text</summary>

Returns the difference of two complex numbers

- inumber1 (number|range, required): Is the complex number from which to subtract inumber2
- inumber2 (number|range, required): Is the complex number to subtract from inumber1

Input: `=IMSUB(10, 10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 14 ms.

</details>

<details>
<summary>IMDIV — exception-like text</summary>

Returns the quotient of two complex numbers

- inumber1 (number|range, required): Is the complex numerator or dividend
- inumber2 (number|range, required): Is the complex denominator or divisor

Input: `=IMDIV(10, 10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 9 ms.

</details>

<details>
<summary>IMPOWER — exception-like text</summary>

Returns a complex number raised to an integer power

- inumber (number|range, required): Is a complex number you want to raise to a power
- numberParam (number|range, required): Is the power to which you want to raise the complex number

Input: `=IMPOWER(10, 10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 9 ms.

</details>

<details>
<summary>IMABS — value returned</summary>

Returns the absolute value (modulus) of a complex number

- inumber (number|range, required): Is a complex number for which you want the absolute value

Input: `=IMABS(10)`

Observed output: `10`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>IMSQRT — exception-like text</summary>

Returns the square root of a complex number

- inumber (number|range, required): Is a complex number for which you want the square root

Input: `=IMSQRT(10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 12 ms.

</details>

<details>
<summary>IMLN — exception-like text</summary>

Returns the natural logarithm of a complex number

- inumber (number|range, required): Is a complex number for which you want the natural logarithm

Input: `=IMLN(10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 11 ms.

</details>

<details>
<summary>IMLOG2 — exception-like text</summary>

Returns the base-2 logarithm of a complex number

- inumber (number|range, required): Is a complex number for which you want the base-2 logarithm

Input: `=IMLOG2(10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 10 ms.

</details>

<details>
<summary>IMLOG10 — exception-like text</summary>

Returns the base-10 logarithm of a complex number

- inumber (number|range, required): Is a complex number for which you want the common logarithm

Input: `=IMLOG10(10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 11 ms.

</details>

<details>
<summary>IMSIN — exception-like text</summary>

Returns the sine of a complex number

- inumber (number|range, required): Is a complex number for which you want the sine

Input: `=IMSIN(10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 8 ms.

</details>

<details>
<summary>IMCOS — exception-like text</summary>

Returns the cosine of a complex number

- inumber (number|range, required): Is a complex number for which you want the cosine

Input: `=IMCOS(10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 8 ms.

</details>

<details>
<summary>IMEXP — exception-like text</summary>

Returns the exponential of a complex number

- inumber (number|range, required): Is a complex number for which you want the exponential

Input: `=IMEXP(10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 11 ms.

</details>

<details>
<summary>IMARGUMENT — value returned</summary>

Returns the argument q, an angle expressed in radians

- inumber (number|range, required): Is a complex number for which you want the argument

Input: `=IMARGUMENT(10)`

Observed output: `0`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>IMCONJUGATE — exception-like text</summary>

Returns the complex conjugate of a complex number

- inumber (number|range, required): Is a complex number for which you want the conjugate

Input: `=IMCONJUGATE(10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 10 ms.

</details>

<details>
<summary>IMAGINARY — value returned</summary>

Returns the imaginary coefficient of a complex number

- inumber (number|range, required): Is a complex number for which you want the imaginary coefficient

Input: `=IMAGINARY(10)`

Observed output: `0`

Classification: value returned. Elapsed: 6 ms.

</details>

<details>
<summary>IMREAL — value returned</summary>

Returns the real coefficient of a complex number

- inumber (number|range, required): Is a complex number for which you want the real coefficient

Input: `=IMREAL(10)`

Observed output: `"10"`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>COMPLEX — Excel error</summary>

Converts real and imaginary coefficients into a complex number

- realNum (value|range, required): Is the real coefficient of the complex number
- iNum (value|range, required): Is the imaginary coefficient of the complex number
- suffix (value|range, optional): Is the suffix for the imaginary component of the complex number

Input: `=COMPLEX(A1, A1, A1)`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 8 ms.

</details>

<details>
<summary>IMSUM — value returned</summary>

Returns the sum of complex numbers

- inumber1 (number|range, required): Are from 1 to 255 complex numbers to add
- inumber2 (number|range, optional): Number or reference.
- rest (value|range, required): Value, reference, or range.

Input: `=IMSUM(10, C1:C5, 10, C1:C5)`

Observed output: `"110"`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>IMPRODUCT — Excel error</summary>

Returns the product of 1 to 255 complex numbers

- inumber1 (number|range, required): Inumber1, Inumber2,... are from 1 to 255 complex numbers to multiply.
- inumber2 (number|range, optional): Number or reference.
- rest (value|range, required): Value, reference, or range.

Input: `=IMPRODUCT(10, C1:C5, 10, C1:C5)`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 10 ms.

</details>

<details>
<summary>DELTA — value returned</summary>

Tests whether two numbers are equal

- number1 (number|range, required): Is the first number
- number2 (number|range, optional): Is the second number

Input: `=DELTA(10, 20)`

Observed output: `0`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>GESTEP — value returned</summary>

Tests whether a number is greater than a threshold value

- numberParam (number|range, required): Is the value to test against step
- step (value|range, optional): Is the threshold value

Input: `=GESTEP(10, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>ERF — value returned</summary>

Returns the error function

- lowerLimit (value|range, required): Is the lower bound for integrating ERF
- upperLimit (value|range, optional): Is the upper bound for integrating ERF

Input: `=ERF(A1, A1)`

Observed output: `0.8427007929497149`

Classification: value returned. Elapsed: 21 ms.

</details>

<details>
<summary>ERFC — value returned</summary>

Returns the complementary error function

- x (value|range, required): Is the lower bound for integrating ERF

Input: `=ERFC(A1)`

Observed output: `0.1572992070502851`

Classification: value returned. Elapsed: 30 ms.

</details>

<details>
<summary>BESSELJ — value returned</summary>

Returns the Bessel function Jn(x)

- x (value|range, required): Is the value at which to evaluate the function
- n (value|range, required): Is the order of the Bessel function

Input: `=BESSELJ(A1, A1)`

Observed output: `0.4400505856771301`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>BESSELK — value returned</summary>

Returns the modified Bessel function Kn(x)

- x (value|range, required): Is the value at which to evaluate the function
- n (value|range, required): Is the order of the function

Input: `=BESSELK(A1, A1)`

Observed output: `0.6019072316669057`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>BESSELY — value returned</summary>

Returns the Bessel function Yn(x)

- x (value|range, required): Is the value at which to evaluate the function
- n (value|range, required): Is the order of the function

Input: `=BESSELY(A1, A1)`

Observed output: `-0.7812128209531197`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>BESSELI — value returned</summary>

Returns the modified Bessel function In(x)

- x (value|range, required): Is the value at which to evaluate the function
- n (value|range, required): Is the order of the Bessel function

Input: `=BESSELI(A1, A1)`

Observed output: `0.5651590975819435`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>CONVERT — exception-like text</summary>

Converts a number from one measurement system to another

- numberParam (number|range, required): Is the value in from_units to convert
- fromUnit (value|range, required): Is the units for number
- toUnit (value|range, required): Is the units for the result

Input: `=CONVERT(10, A1, A1)`

Observed output: `"t.substring is not a function"`

Classification: exception-like text. Elapsed: 10 ms.

</details>

<details>
<summary>ERF.PRECISE — value returned</summary>

Returns the error function

- x (value|range, required): Is the lower bound for integrating ERF.PRECISE

Input: `=ERF.PRECISE(A1)`

Observed output: `0.8427007929497149`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>ERFC.PRECISE — value returned</summary>

Returns the complementary error function

- x (value|range, required): Is the lower bound for integrating ERFC.PRECISE

Input: `=ERFC.PRECISE(A1)`

Observed output: `0.1572992070502851`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>IMTAN — exception-like text</summary>

Returns the tangent of a complex number

- inumber (number|range, required): Is a complex number for which you want the tangent

Input: `=IMTAN(10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 7 ms.

</details>

<details>
<summary>IMCOT — exception-like text</summary>

Returns the cotangent of a complex number

- inumber (number|range, required): Is a complex number for which you want the cotangent

Input: `=IMCOT(10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 10 ms.

</details>

<details>
<summary>IMCSC — exception-like text</summary>

Returns the cosecant of a complex number

- inumber (number|range, required): Is a complex number for which you want the cosecant

Input: `=IMCSC(10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 13 ms.

</details>

<details>
<summary>IMCSCH — exception-like text</summary>

Returns the hyperbolic cosecant of a complex number

- inumber (number|range, required): Is a complex number for which you want the hyperbolic cosecant

Input: `=IMCSCH(10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 9 ms.

</details>

<details>
<summary>IMSEC — exception-like text</summary>

Returns the secant of a complex number

- inumber (number|range, required): Is a complex number for which you want the secant

Input: `=IMSEC(10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 14 ms.

</details>

<details>
<summary>IMSECH — exception-like text</summary>

Returns the hyperbolic secant of a complex number

- inumber (number|range, required): Is a complex number for which you want the hyperbolic secant

Input: `=IMSECH(10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 14 ms.

</details>

<details>
<summary>BITAND — value returned</summary>

Returns a bitwise 'And' of two numbers

- number1 (number|range, required): Is the decimal representation of the binary number you want to evaluate
- number2 (number|range, required): Is the decimal representation of the binary number you want to evaluate

Input: `=BITAND(10, 20)`

Observed output: `0`

Classification: value returned. Elapsed: 6 ms.

</details>

<details>
<summary>BITOR — value returned</summary>

Returns a bitwise 'Or' of two numbers

- number1 (number|range, required): Is the decimal representation of the binary number you want to evaluate
- number2 (number|range, required): Is the decimal representation of the binary number you want to evaluate

Input: `=BITOR(10, 20)`

Observed output: `30`

Classification: value returned. Elapsed: 6 ms.

</details>

<details>
<summary>BITXOR — value returned</summary>

Returns a bitwise 'Exclusive Or' of two numbers

- number1 (number|range, required): Is the decimal representation of the binary number you want to evaluate
- number2 (number|range, required): Is the decimal representation of the binary number you want to evaluate

Input: `=BITXOR(10, 20)`

Observed output: `30`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>BITLSHIFT — value returned</summary>

Returns a number shifted left by shift_amount bits

- numberParam (number|range, required): Is the decimal representation of the binary number you want to evaluate
- shiftAmount (value|range, required): Is the number of bits that you want to shift Number left by

Input: `=BITLSHIFT(10, A1)`

Observed output: `20`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>BITRSHIFT — value returned</summary>

Returns a number shifted right by shift_amount bits

- numberParam (number|range, required): Is the decimal representation of the binary number you want to evaluate
- shiftAmount (value|range, required): Is the number of bits that you want to shift Number right by

Input: `=BITRSHIFT(10, A1)`

Observed output: `5`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>IMSINH — exception-like text</summary>

Returns the hyperbolic sine of a complex number

- inumber (number|range, required): Is a complex number for which you want the hyperbolic sine

Input: `=IMSINH(10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 13 ms.

</details>

<details>
<summary>IMCOSH — exception-like text</summary>

Returns the hyperbolic cosine of a complex number

- inumber (number|range, required): Is a complex number for which you want the hyperbolic cosine

Input: `=IMCOSH(10)`

Observed output: `"e.substring is not a function"`

Classification: exception-like text. Elapsed: 9 ms.

</details>


### Formula category: financial

<details>
<summary>NPV — value returned</summary>

Returns the net present value of an investment based on a discount rate and a series of future payments (negative values) and income (positive values)

- rate (value|range, required): Is the rate of discount over the length of one period
- value1 (value|range, required): Are 1 to 254 payments and income, equally spaced in time and occurring at the end of each period
- value2 (value|range, optional): Value argument for the calculation.
- rest (value|range, required): Value, reference, or range.

Input: `=NPV(A1, A1:A5, C1:C5, C1:C5)`

Observed output: `1.953460693359375`

Classification: value returned. Elapsed: 16 ms.

</details>

<details>
<summary>PV — value returned</summary>

Returns the present value of an investment: the total amount that a series of future payments is worth now

- rate (value|range, required): Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR
- nper (value|range, required): Is the total number of payment periods in an investment
- pmt (value|range, required): Is the payment made each period and cannot change over the life of the investment
- fv (value|range, optional): Is the future value, or a cash balance you want to attain after the last payment is made
- typeParam (value|range, optional): Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted

Input: `=PV(A1, A1, A1)`

Observed output: `-0.5`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>FV — value returned</summary>

Returns the future value of an investment based on periodic, constant payments and a constant interest rate

- rate (value|range, required): Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR
- nper (value|range, required): Is the total number of payment periods in the investment
- pmt (value|range, required): Is the payment made each period; it cannot change over the life of the investment
- pv (value|range, optional): Is the present value, or the lump-sum amount that a series of future payments is worth now. If omitted, Pv = 0
- typeParam (value|range, optional): Is a value representing the timing of payment: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted

Input: `=FV(A1, A1, A1)`

Observed output: `-1`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>NPER — value returned</summary>

Returns the number of periods for an investment based on periodic, constant payments and a constant interest rate

- rate (value|range, required): Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR
- pmt (value|range, required): Is the payment made each period; it cannot change over the life of the investment
- pv (value|range, required): Is the present value, or the lump-sum amount that a series of future payments is worth now
- fv (value|range, optional): Is the future value, or a cash balance you want to attain after the last payment is made. If omitted, zero is used
- typeParam (value|range, optional): Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted

Input: `=NPER(A1, A1, A1)`

Observed output: `-1`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>PMT — value returned</summary>

Calculates the payment for a loan based on constant payments and a constant interest rate

- rate (value|range, required): Is the interest rate per period for the loan. For example, use 6%/4 for quarterly payments at 6% APR
- nper (value|range, required): Is the total number of payments for the loan
- pv (value|range, required): Is the present value: the total amount that a series of future payments is worth now
- fv (value|range, optional): Is the future value, or a cash balance you want to attain after the last payment is made, 0 (zero) if omitted
- typeParam (value|range, optional): Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted

Input: `=PMT(A1, A1, A1)`

Observed output: `-2`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>RATE — Excel error</summary>

Returns the interest rate per period of a loan or an investment. For example, use 6%/4 for quarterly payments at 6% APR

- nper (value|range, required): Is the total number of payment periods for the loan or investment
- pmt (value|range, required): Is the payment made each period and cannot change over the life of the loan or investment
- pv (value|range, required): Is the present value: the total amount that a series of future payments is worth now
- fv (value|range, optional): Is the future value, or a cash balance you want to attain after the last payment is made. If omitted, uses Fv = 0
- typeParam (value|range, optional): Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted
- guess (value|range, optional): Is your guess for what the rate will be; if omitted, Guess = 0.1 (10 percent)

Input: `=RATE(A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 12 ms.

</details>

<details>
<summary>MIRR — Excel error</summary>

Returns the internal rate of return for a series of periodic cash flows, considering both cost of investment and interest on reinvestment of cash

- values (value|range, required): Is an array or a reference to cells that contain numbers that represent a series of payments (negative) and income (positive) at regular periods
- financeRate (value|range, required): Is the interest rate you pay on the money used in the cash flows
- reinvestRate (value|range, required): Is the interest rate you receive on the cash flows as you reinvest them

Input: `=MIRR(A1:A5, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 7 ms.

</details>

<details>
<summary>IRR — Excel error</summary>

Returns the internal rate of return for a series of cash flows

- values (value|range, required): Is an array or a reference to cells that contain numbers for which you want to calculate the internal rate of return
- guess (value|range, optional): Is a number that you guess is close to the result of IRR; 0.1 (10 percent) if omitted

Input: `=IRR(A1:A5, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 8 ms.

</details>

<details>
<summary>SLN — value returned</summary>

Returns the straight-line depreciation of an asset for one period

- cost (value|range, required): Is the initial cost of the asset
- salvage (value|range, required): Is the salvage value at the end of the life of the asset
- life (value|range, required): Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset)

Input: `=SLN(A1, A1, A1)`

Observed output: `0`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>SYD — value returned</summary>

Returns the sum-of-years' digits depreciation of an asset for a specified period

- cost (value|range, required): Is the initial cost of the asset
- salvage (value|range, required): Is the salvage value at the end of the life of the asset
- life (value|range, required): Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset)
- per (value|range, required): Is the period and must use the same units as Life

Input: `=SYD(A1, A1, A1, A1)`

Observed output: `0`

Classification: value returned. Elapsed: 15 ms.

</details>

<details>
<summary>DDB — value returned</summary>

Returns the depreciation of an asset for a specified period using the double-declining balance method or some other method you specify

- cost (value|range, required): Is the initial cost of the asset
- salvage (value|range, required): Is the salvage value at the end of the life of the asset
- life (value|range, required): Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset)
- period (value|range, required): Is the period for which you want to calculate the depreciation. Period must use the same units as Life
- factor (value|range, optional): Is the rate at which the balance declines. If Factor is omitted, it is assumed to be 2 (the double-declining balance method)

Input: `=DDB(A1, A1, A1, A1)`

Observed output: `0`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>IPMT — value returned</summary>

Returns the interest payment for a given period for an investment, based on periodic, constant payments and a constant interest rate

- rate (value|range, required): Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR
- per (value|range, required): Is the period for which you want to find the interest and must be in the range 1 to Nper
- nper (value|range, required): Is the total number of payment periods in an investment
- pv (value|range, required): Is the present value, or the lump-sum amount that a series of future payments is worth now
- fv (value|range, optional): Is the future value, or a cash balance you want to attain after the last payment is made. If omitted, Fv = 0
- typeParam (value|range, optional): Is a logical value representing the timing of payment: at the end of the period = 0 or omitted, at the beginning of the period = 1

Input: `=IPMT(A1, A1, A1, A1)`

Observed output: `-1`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>PPMT — value returned</summary>

Returns the payment on the principal for a given investment based on periodic, constant payments and a constant interest rate

- rate (value|range, required): Is the interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR
- per (value|range, required): Specifies the period and must be in the range 1 to nper
- nper (value|range, required): Is the total number of payment periods in an investment
- pv (value|range, required): Is the present value: the total amount that a series of future payments is worth now
- fv (value|range, optional): Is the future value, or cash balance you want to attain after the last payment is made
- typeParam (value|range, optional): Is a logical value: payment at the beginning of the period = 1; payment at the end of the period = 0 or omitted

Input: `=PPMT(A1, A1, A1, A1)`

Observed output: `-1`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>VDB — unimplemented text</summary>

Returns the depreciation of an asset for any period you specify, including partial periods, using the double-declining balance method or some other method you specify

- cost (value|range, required): Is the initial cost of the asset
- salvage (value|range, required): Is the salvage value at the end of the life of the asset
- life (value|range, required): Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset)
- startPeriod (value|range, required): Is the starting period for which you want to calculate the depreciation, in the same units as Life
- endPeriod (value|range, required): Is the ending period for which you want to calculate the depreciation, in the same units as Life
- factor (value|range, optional): Is the rate at which the balance declines, 2 (double-declining balance) if omitted
- noSwitch (value|range, optional): Switch to straight-line depreciation when depreciation is greater than the declining balance = FALSE or omitted; do not switch = TRUE

Input: `=VDB(A1, A1, A1, A1)`

Observed output: `"VDB is not implemented. cost=1, salvage=1, life=1, startPeriod=1, endPeriod=undefined, factor=undefined, noSwitch=undefined"`

Classification: unimplemented text. Elapsed: 10 ms.

</details>

<details>
<summary>DB — value returned</summary>

Returns the depreciation of an asset for a specified period using the fixed-declining balance method

- cost (value|range, required): Is the initial cost of the asset
- salvage (value|range, required): Is the salvage value at the end of the life of the asset
- life (value|range, required): Is the number of periods over which the asset is being depreciated (sometimes called the useful life of the asset)
- period (value|range, required): Is the period for which you want to calculate the depreciation. Period must use the same units as Life
- month (value|range, optional): Is the number of months in the first year. If month is omitted, it is assumed to be 12

Input: `=DB(A1, A1, A1, A1)`

Observed output: `0`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>ISPMT — value returned</summary>

Returns the interest paid during a specific period of an investment

- rate (value|range, required): Interest rate per period. For example, use 6%/4 for quarterly payments at 6% APR
- per (value|range, required): Period for which you want to find the interest
- nper (value|range, required): Number of payment periods in an investment
- pv (value|range, required): Lump sum amount that a series of future payments is right now

Input: `=ISPMT(A1, A1, A1, A1)`

Observed output: `0`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>XIRR — Excel error</summary>

Returns the internal rate of return for a schedule of cash flows

- values (value|range, required): Is a series of cash flows that correspond to a schedule of payments in dates
- dates (date, required): Is a schedule of payment dates that corresponds to the cash flow payments
- guess (value|range, optional): Is a number that you guess is close to the result of XIRR

Input: `=XIRR(A1:A5, DATE(2024, 1, 1), A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 13 ms.

</details>

<details>
<summary>XNPV — Excel error</summary>

Returns the net present value for a schedule of cash flows

- rate (value|range, required): Is the discount rate to apply to the cash flows
- values (value|range, required): Is a series of cash flows that correspond to a schedule of payments in dates
- dates (date, required): Is a schedule of payment dates that corresponds to the cash flow payments

Input: `=XNPV(A1, A1:A5, DATE(2024, 1, 1))`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 10 ms.

</details>

<details>
<summary>PRICEMAT — Excel error</summary>

Returns the price per $100 face value of a security that pays interest at maturity

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- issue (boolean, required): Is the security's issue date, expressed as a serial date number
- rate (value|range, required): Is the security's interest rate at date of issue
- yld (value|range, required): Is the security's annual yield
- basis (value|range, optional): Is the type of day count basis to use

Input: `=PRICEMAT(A1, A1, TRUE, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 10 ms.

</details>

<details>
<summary>YIELDMAT — Excel error</summary>

Returns the annual yield of a security that pays interest at maturity

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- issue (boolean, required): Is the security's issue date, expressed as a serial date number
- rate (value|range, required): Is the security's interest rate at date of issue
- pr (value|range, required): Is the security's price per $100 face value
- basis (value|range, optional): Is the type of day count basis to use

Input: `=YIELDMAT(A1, A1, TRUE, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 12 ms.

</details>

<details>
<summary>INTRATE — unimplemented text</summary>

Returns the interest rate for a fully invested security

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- investment (value|range, required): Is the amount invested in the security
- redemption (value|range, required): Is the amount to be received at maturity
- basis (value|range, optional): Is the type of day count basis to use

Input: `=INTRATE(A1, A1, A1, A1)`

Observed output: `"INTRATE is not implemented. settlement=1, maturity=1, investment=1, redemption=1, basis=undefined"`

Classification: unimplemented text. Elapsed: 14 ms.

</details>

<details>
<summary>RECEIVED — Excel error</summary>

Returns the amount received at maturity for a fully invested security

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- investment (value|range, required): Is the amount invested in the security
- discount (value|range, required): Is the security's discount rate
- basis (value|range, optional): Is the type of day count basis to use

Input: `=RECEIVED(A1, A1, A1, 10)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 20 ms.

</details>

<details>
<summary>DISC — Excel error</summary>

Returns the discount rate for a security

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- pr (value|range, required): Is the security's price per $100 face value
- redemption (value|range, required): Is the security's redemption value per $100 face value
- basis (value|range, optional): Is the type of day count basis to use

Input: `=DISC(A1, A1, A1, A1)`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 9 ms.

</details>

<details>
<summary>PRICEDISC — Excel error</summary>

Returns the price per $100 face value of a discounted security

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- discount (value|range, required): Is the security's discount rate
- redemption (value|range, required): Is the security's redemption value per $100 face value
- basis (value|range, optional): Is the type of day count basis to use

Input: `=PRICEDISC(A1, A1, 10, A1)`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 8 ms.

</details>

<details>
<summary>YIELDDISC — Excel error</summary>

Returns the annual yield for a discounted security. For example, a treasury bill

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- pr (value|range, required): Is the security's price per $100 face value
- redemption (value|range, required): Is the security's redemption value per $100 face value
- basis (value|range, optional): Is the type of day count basis to use

Input: `=YIELDDISC(A1, A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 8 ms.

</details>

<details>
<summary>TBILLEQ — value returned</summary>

Returns the bond-equivalent yield for a treasury bill

- settlement (value|range, required): Is the Treasury bill's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the Treasury bill's maturity date, expressed as a serial date number
- discount (value|range, required): Is the Treasury bill's discount rate

Input: `=TBILLEQ(A1, A1, 10)`

Observed output: `10.13888888888889`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>TBILLPRICE — value returned</summary>

Returns the price per $100 face value for a treasury bill

- settlement (value|range, required): Is the Treasury bill's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the Treasury bill's maturity date, expressed as a serial date number
- discount (value|range, required): Is the Treasury bill's discount rate

Input: `=TBILLPRICE(A1, A1, 10)`

Observed output: `100`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>TBILLYIELD — Excel error</summary>

Returns the yield for a treasury bill

- settlement (value|range, required): Is the Treasury bill's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the Treasury bill's maturity date, expressed as a serial date number
- pr (value|range, required): Is the Treasury Bill's price per $100 face value

Input: `=TBILLYIELD(A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 7 ms.

</details>

<details>
<summary>PRICE — Excel error</summary>

Returns the price per $100 face value of a security that pays periodic interest

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- rate (value|range, required): Is the security's annual coupon rate
- yld (value|range, required): Is the security's annual yield
- redemption (value|range, required): Is the security's redemption value per $100 face value
- frequency (value|range, required): Is the number of coupon payments per year
- basis (value|range, optional): Is the type of day count basis to use

Input: `=PRICE(A1, A1, A1, A1)`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 14 ms.

</details>

<details>
<summary>YIELD — Excel error</summary>

Returns the yield on a security that pays periodic interest

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- rate (value|range, required): Is the security's annual coupon rate
- pr (value|range, required): Is the security's price per $100 face value
- redemption (value|range, required): Is the security's redemption value per $100 face value
- frequency (value|range, required): Is the number of coupon payments per year
- basis (value|range, optional): Is the type of day count basis to use

Input: `=YIELD(A1, A1, A1, A1)`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 14 ms.

</details>

<details>
<summary>DOLLARDE — value returned</summary>

Converts a dollar price, expressed as a fraction, into a dollar price, expressed as a decimal number

- fractionalDollar (value|range, required): Is a number expressed as a fraction
- fraction (value|range, required): Is the integer to use in the denominator of the fraction

Input: `=DOLLARDE(A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>DOLLARFR — value returned</summary>

Converts a dollar price, expressed as a decimal number, into a dollar price, expressed as a fraction

- decimalDollar (value|range, required): Is a decimal number
- fraction (value|range, required): Is the integer to use in the denominator of a fraction

Input: `=DOLLARFR(A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>NOMINAL — value returned</summary>

Returns the annual nominal interest rate

- effectRate (value|range, required): Is the effective interest rate
- npery (value|range, required): Is the number of compounding periods per year

Input: `=NOMINAL(A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>EFFECT — value returned</summary>

Returns the effective annual interest rate

- nominalRate (value|range, required): Is the nominal interest rate
- npery (value|range, required): Is the number of compounding periods per year

Input: `=EFFECT(A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>CUMPRINC — Excel error</summary>

Returns the cumulative principal paid on a loan between two periods

- rate (value|range, required): Is the interest rate
- nper (value|range, required): Is the total number of payment periods
- pv (value|range, required): Is the present value
- startPeriod (value|range, required): Is the first period in the calculation
- endPeriod (value|range, required): Is the last period in the calculation
- typeParam (value|range, required): Is the timing of the payment

Input: `=CUMPRINC(A1, A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 14 ms.

</details>

<details>
<summary>CUMIPMT — Excel error</summary>

Returns the cumulative interest paid between two periods

- rate (value|range, required): Is the interest rate
- nper (value|range, required): Is the total number of payment periods
- pv (value|range, required): Is the present value
- startPeriod (value|range, required): Is the first period in the calculation
- endPeriod (value|range, required): Is the last period in the calculation
- typeParam (value|range, required): Is the timing of the payment

Input: `=CUMIPMT(A1, A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 11 ms.

</details>

<details>
<summary>COUPDAYBS — Excel error</summary>

Returns the number of days from the beginning of the coupon period to the settlement date

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- frequency (value|range, required): Is the number of coupon payments per year
- basis (value|range, optional): Is the type of day count basis to use

Input: `=COUPDAYBS(A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 9 ms.

</details>

<details>
<summary>COUPDAYS — Excel error</summary>

Returns the number of days in the coupon period that contains the settlement date

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- frequency (value|range, required): Is the number of coupon payments per year
- basis (value|range, optional): Is the type of day count basis to use

Input: `=COUPDAYS(A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 13 ms.

</details>

<details>
<summary>COUPDAYSNC — Excel error</summary>

Returns the number of days from the settlement date to the next coupon date

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- frequency (value|range, required): Is the number of coupon payments per year
- basis (value|range, optional): Is the type of day count basis to use

Input: `=COUPDAYSNC(A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 10 ms.

</details>

<details>
<summary>COUPNCD — Excel error</summary>

Returns the next coupon date after the settlement date

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- frequency (value|range, required): Is the number of coupon payments per year
- basis (value|range, optional): Is the type of day count basis to use

Input: `=COUPNCD(A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 12 ms.

</details>

<details>
<summary>COUPNUM — Excel error</summary>

Returns the number of coupons payable between the settlement date and maturity date

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- frequency (value|range, required): Is the number of coupon payments per year
- basis (value|range, optional): Is the type of day count basis to use

Input: `=COUPNUM(A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 12 ms.

</details>

<details>
<summary>COUPPCD — Excel error</summary>

Returns the previous coupon date before the settlement date

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- frequency (value|range, required): Is the number of coupon payments per year
- basis (value|range, optional): Is the type of day count basis to use

Input: `=COUPPCD(A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 8 ms.

</details>

<details>
<summary>DURATION — unimplemented text</summary>

Returns the annual duration of a security with periodic interest payments

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- coupon (value|range, required): Is the security's annual coupon rate
- yld (value|range, required): Is the security's annual yield
- frequency (value|range, required): Is the number of coupon payments per year
- basis (value|range, optional): Is the type of day count basis to use

Input: `=DURATION(A1, A1, A1, A1)`

Observed output: `"DURATION is not implemented. settlement=1, maturity=1, coupon=1, yld=1, frequency=undefined, basis=undefined"`

Classification: unimplemented text. Elapsed: 9 ms.

</details>

<details>
<summary>MDURATION — unimplemented text</summary>

Returns the Macauley modified duration for a security with an assumed par value of $100

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- coupon (value|range, required): Is the security's annual coupon rate
- yld (value|range, required): Is the security's annual yield
- frequency (value|range, required): Is the number of coupon payments per year
- basis (value|range, optional): Is the type of day count basis to use

Input: `=MDURATION(A1, A1, A1, A1)`

Observed output: `"MDURATION is not implemented. settlement=1, maturity=1, coupon=1, yld=1, frequency=undefined, basis=undefined"`

Classification: unimplemented text. Elapsed: 7 ms.

</details>

<details>
<summary>ODDLPRICE — unimplemented text</summary>

Returns the price per $100 face value of a security with an odd last period

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- lastInterest (value|range, required): Is the security's last coupon date, expressed as a serial date number
- rate (value|range, required): Is the security's interest rate
- yld (value|range, required): Is the security's annual yield
- redemption (value|range, required): Is the security's redemption value per $100 face value
- frequency (value|range, required): Is the number of coupon payments per year
- basis (value|range, optional): Is the type of day count basis to use

Input: `=ODDLPRICE(A1, A1, A1, A1)`

Observed output: `"ODDLPRICE is not implemented. settlement=1, maturity=1, lastInterest=1, rate=1, yld=undefined, redemption=undefined, frequency=undefined, basis=undefined"`

Classification: unimplemented text. Elapsed: 7 ms.

</details>

<details>
<summary>ODDLYIELD — unimplemented text</summary>

Returns the yield of a security with an odd last period

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- lastInterest (value|range, required): Is the security's last coupon date, expressed as a serial date number
- rate (value|range, required): Is the security's interest rate
- pr (value|range, required): Is the security's price
- redemption (value|range, required): Is the security's redemption value per $100 face value
- frequency (value|range, required): Is the number of coupon payments per year
- basis (value|range, optional): Is the type of day count basis to use

Input: `=ODDLYIELD(A1, A1, A1, A1)`

Observed output: `"ODDLYIELD is not implemented. settlement=1, maturity=1, lastInterest=1, rate=1, pr=undefined, redemption=undefined, frequency=undefined, basis=undefined"`

Classification: unimplemented text. Elapsed: 7 ms.

</details>

<details>
<summary>ODDFPRICE — unimplemented text</summary>

Returns the price per $100 face value of a security with an odd first period

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- issue (boolean, required): Is the security's issue date, expressed as a serial date number
- firstCoupon (value|range, required): Is the security's first coupon date, expressed as a serial date number
- rate (value|range, required): Is the security's interest rate
- yld (value|range, required): Is the security's annual yield
- redemption (value|range, required): Is the security's redemption value per $100 face value
- frequency (value|range, required): Is the number of coupon payments per year
- basis (value|range, optional): Is the type of day count basis to use

Input: `=ODDFPRICE(A1, A1, TRUE, A1)`

Observed output: `"ODDFPRICE is not implemented. settlement=1, maturity=1, issue=true, firstCoupon=1, rate=undefined, yld=undefined, redemption=undefined, frequency=undefined, basis=undefined"`

Classification: unimplemented text. Elapsed: 8 ms.

</details>

<details>
<summary>ODDFYIELD — unimplemented text</summary>

Returns the yield of a security with an odd first period

- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- maturity (value|range, required): Is the security's maturity date, expressed as a serial date number
- issue (boolean, required): Is the security's issue date, expressed as a serial date number
- firstCoupon (value|range, required): Is the security's first coupon date, expressed as a serial date number
- rate (value|range, required): Is the security's interest rate
- pr (value|range, required): Is the security's price
- redemption (value|range, required): Is the security's redemption value per $100 face value
- frequency (value|range, required): Is the number of coupon payments per year
- basis (value|range, optional): Is the type of day count basis to use

Input: `=ODDFYIELD(A1, A1, TRUE, A1)`

Observed output: `"ODDFYIELD is not implemented. settlement=1, maturity=1, issue=true, firstCoupon=1, rate=undefined, pr=undefined, redemption=undefined, frequency=undefined, basis=undefined"`

Classification: unimplemented text. Elapsed: 11 ms.

</details>

<details>
<summary>AMORLINC — unimplemented text</summary>

Returns the prorated linear depreciation of an asset for each accounting period.

- cost (value|range, required): Is the cost of the asset
- datePurchased (date, required): Is the date the asset is purchased
- firstPeriod (value|range, required): Is the date of the end of the first period
- salvage (value|range, required): Is the salvage value at the end of life of the asset.
- period (value|range, required): Is the period
- rate (value|range, required): Is the rate of depreciation
- basis (value|range, optional): Year_basis : 0 for year of 360 days, 1 for actual, 3 for year of 365 days.

Input: `=AMORLINC(A1, DATE(2024, 1, 1), A1, A1)`

Observed output: `"AMORLINC is not implemented. cost=1, datePurchased=45292, firstPeriod=1, salvage=1, period=undefined, rate=undefined, basis=undefined"`

Classification: unimplemented text. Elapsed: 11 ms.

</details>

<details>
<summary>ACCRINT — Excel error</summary>

Returns the accrued interest for a security that pays periodic interest.

- issue (boolean, required): Is the security's issue date, expressed as a serial date number
- firstInterest (value|range, required): Is the security's first interest date, expressed as a serial date number
- settlement (value|range, required): Is the security's settlement date, expressed as a serial date number
- rate (value|range, required): Is the security's annual coupon rate
- par (value|range, required): Is the security's par value
- frequency (value|range, required): Is the number of coupon payments per year
- basis (value|range, optional): Is the type of day count basis to use
- _calcMethod (value|range, optional): Value, reference, or range.

Input: `=ACCRINT(TRUE, A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 11 ms.

</details>

<details>
<summary>ACCRINTM — unimplemented text</summary>

Returns the accrued interest for a security that pays interest at maturity

- issue (boolean, required): Is the security's issue date, expressed as a serial date number
- settlement (value|range, required): Is the security's maturity date, expressed as a serial date number
- rate (value|range, required): Is the security's annual coupon rate
- par (value|range, required): Is the security's par value
- basis (value|range, optional): Is the type of day count basis to use

Input: `=ACCRINTM(TRUE, A1, A1, A1)`

Observed output: `"ACCRINTM is not implemented. issue=true, settlement=1, rate=1, par=1, basis=undefined"`

Classification: unimplemented text. Elapsed: 10 ms.

</details>

<details>
<summary>FVSCHEDULE — value returned</summary>

Returns the future value of an initial principal after applying a series of compound interest rates

- principal (value|range, required): Is the present value
- schedule (value|range, required): Is an array of interest rates to apply

Input: `=FVSCHEDULE(A1, A1)`

Observed output: `2`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>PDURATION — value returned</summary>

Returns the number of periods required by an investment to reach a specified value

- rate (value|range, required): Is the interest rate per period.
- pv (value|range, required): Is the present value of the investment
- fv (value|range, required): Is the desired future value of the investment

Input: `=PDURATION(A1, A1, A1)`

Observed output: `0`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>RRI — value returned</summary>

Returns an equivalent interest rate for the growth of an investment

- nper (value|range, required): Is the number of periods for the investment
- pv (value|range, required): Is the present value of the investment
- fv (value|range, required): Is the future value of the investment

Input: `=RRI(A1, A1, A1)`

Observed output: `0`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>STOCKHISTORY — unimplemented text</summary>

Returns an array of historical quote data for a symbol and date range you specify.

- stock (value|range, required): Symbol of financial instrument to be considered or a Stock data type.
- startDate (date, required): First date to return data from.
- endDate (date, optional): Last date to return data from.
- interval (value|range, optional): A number indicating the granularity of the data; 0 - Daily, 1 - Weekly, 2 - Monthly
- headers (value|range, optional): A logical value to add column header data; 0 - No column header, 1 - Show column header, 2 - Show instrument identifier and column header
- properties1 (value|range, optional): A number indicating which column of data to return; 0 through 5
- rest (value|range, required): Value, reference, or range.

Input: `=STOCKHISTORY(A1, DATE(2024, 1, 1), C1:C5, C1:C5)`

Observed output: `"STOCKHISTORY is not implemented. stock=1, startDate=45292, endDate=3,6,9,12,15, interval=3,6,9,12,15, headers=undefined, properties1=undefined, rest.length=0"`

Classification: unimplemented text. Elapsed: 12 ms.

</details>


### Formula category: information

<details>
<summary>ISNA — value returned</summary>

Checks whether a value is #N/A, and returns TRUE or FALSE

- value (value|range, required): Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value

Input: `=ISNA(A1:A5)`

Observed output: `false`

Classification: value returned. Elapsed: 7 ms.

</details>

<details>
<summary>ISERROR — value returned</summary>

Checks whether a value is an error, and returns TRUE or FALSE

- value (value|range, required): Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value

Input: `=ISERROR(A1:A5)`

Observed output: `false`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>NA — Excel error</summary>

Returns the error value #N/A (value not available)


Input: `=NA()`

Observed output: `"#N/A"`

Classification: Excel error. Elapsed: 7 ms.

</details>

<details>
<summary>TYPE — value returned</summary>

Returns an integer representing the data type of a value: number = 1; text = 2; logical value = 4; error value = 16; array = 64; compound data = 128

- value (value|range, required): Can be any value

Input: `=TYPE(A1:A5)`

Observed output: `1`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>ISREF — unimplemented text</summary>

Checks whether a value is a reference, and returns TRUE or FALSE

- value (value|range, required): Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value

Input: `=ISREF(A1:A5)`

Observed output: `"ISREF is not implemented. value=1,2,3,4,5"`

Classification: unimplemented text. Elapsed: 10 ms.

</details>

<details>
<summary>ISERR — value returned</summary>

Checks whether a value is an error other than #N/A, and returns TRUE or FALSE

- value (value|range, required): Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value

Input: `=ISERR(A1:A5)`

Observed output: `false`

Classification: value returned. Elapsed: 16 ms.

</details>

<details>
<summary>ISTEXT — value returned</summary>

Checks whether a value is text, and returns TRUE or FALSE

- value (value|range, required): Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value

Input: `=ISTEXT(A1:A5)`

Observed output: `false`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>ISNUMBER — value returned</summary>

Checks whether a value is a number, and returns TRUE or FALSE

- value (value|range, required): Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value

Input: `=ISNUMBER(A1:A5)`

Observed output: `false`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>ISBLANK — value returned</summary>

Checks whether a reference is to an empty cell, and returns TRUE or FALSE

- value (value|range, required): Is the cell or a name that refers to the cell you want to test

Input: `=ISBLANK(A1:A5)`

Observed output: `false`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>N — value returned</summary>

Converts non-number value to a number, dates to serial numbers, TRUE to 1, anything else to 0 (zero)

- value (value|range, required): Is the value you want converted

Input: `=N(A1:A5)`

Observed output: `0`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>ISNONTEXT — value returned</summary>

Checks whether a value is not text (blank cells are not text), and returns TRUE or FALSE

- value (value|range, required): Is the value you want tested: a cell; a formula; or a name referring to a cell, formula, or value

Input: `=ISNONTEXT(A1:A5)`

Observed output: `true`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>ISLOGICAL — value returned</summary>

Checks whether a value is a logical value (TRUE or FALSE), and returns TRUE or FALSE

- value (value|range, required): Is the value you want to test. Value can refer to a cell, a formula, or a name that refers to a cell, formula, or value

Input: `=ISLOGICAL(A1:A5)`

Observed output: `false`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>ERROR.TYPE — Excel error</summary>

Returns a number matching an error value.

- errorVal (value|range, required): Is the error value for which you want the identifying number, and can be an actual error value or a reference to a cell containing an error value

Input: `=ERROR.TYPE(A1)`

Observed output: `"#N/A"`

Classification: Excel error. Elapsed: 8 ms.

</details>

<details>
<summary>ISEVEN — value returned</summary>

Returns TRUE if the number is even

- numberParam (number|range, required): Is the value to test

Input: `=ISEVEN(10)`

Observed output: `true`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>ISODD — value returned</summary>

Returns TRUE if the number is odd

- numberParam (number|range, required): Is the value to test

Input: `=ISODD(10)`

Observed output: `false`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>SHEET — value returned</summary>

Returns the sheet number of the referenced sheet

- value (value|range, optional): Is the name of a sheet or a reference that you want the sheet number of. If omitted the number of the sheet containing the function is returned

Input: `=SHEET(A1:A5)`

Observed output: `1`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>SHEETS — value returned</summary>

Returns the number of sheets in a reference

- reference (range, optional): Is a reference for which you want to know the number of sheets it contains. If omitted the number of sheets in the workbook containing the function is returned

Input: `=SHEETS(A1:A5)`

Observed output: `1`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>ISFORMULA — unimplemented text</summary>

Checks whether a reference is to a cell containing a formula, and returns TRUE or FALSE

- reference (range, required): Is a reference to the cell you want to test. Reference can be a cell reference, a formula, or name that refers to a cell

Input: `=ISFORMULA(A1:A5)`

Observed output: `"ISFORMULA is not implemented. reference=1,2,3,4,5"`

Classification: unimplemented text. Elapsed: 15 ms.

</details>

<details>
<summary>ISOMITTED — unimplemented text</summary>

Checks whether the value is omitted, and returns TRUE or FALSE

- argument (value|range, required): Is the value you want to test, such as a LAMBDA parameter

Input: `=ISOMITTED(A1)`

Observed output: `"ISOMITTED is not implemented. argument=1"`

Classification: unimplemented text. Elapsed: 7 ms.

</details>


### Formula category: logical

<details>
<summary>IF — value returned</summary>

Checks whether a condition is met, and returns one value if TRUE, and another value if FALSE

- logicalTest (boolean, required): Is any value or expression that can be evaluated to TRUE or FALSE
- valueIfTrue (value|range, optional): Is the value that is returned if Logical_test is TRUE. If omitted, TRUE is returned. You can nest up to seven IF functions
- valueIfFalse (value|range, optional): Is the value that is returned if Logical_test is FALSE. If omitted, FALSE is returned

Input: `=IF(TRUE, A1:A5)`

Observed output: `1`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>TRUE — value returned</summary>

Returns the logical value TRUE


Input: `=TRUE()`

Observed output: `true`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>FALSE — value returned</summary>

Returns the logical value FALSE


Input: `=FALSE()`

Observed output: `false`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>AND — value returned</summary>

Checks whether all arguments are TRUE, and returns TRUE if all arguments are TRUE

- logical1 (boolean, required): Are 1 to 255 conditions you want to test that can be either TRUE or FALSE and can be logical values, arrays, or references
- logical2 (boolean, optional): TRUE or FALSE.
- rest (value|range, required): Value, reference, or range.

Input: `=AND(TRUE, C1:C5, TRUE, C1:C5)`

Observed output: `true`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>OR — value returned</summary>

Checks whether any of the arguments are TRUE, and returns TRUE or FALSE. Returns FALSE only if all arguments are FALSE

- logical1 (boolean, required): Are 1 to 255 conditions that you want to test that can be either TRUE or FALSE
- logical2 (boolean, optional): TRUE or FALSE.
- rest (value|range, required): Value, reference, or range.

Input: `=OR(TRUE, C1:C5, TRUE, C1:C5)`

Observed output: `true`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>NOT — value returned</summary>

Changes FALSE to TRUE, or TRUE to FALSE

- logical (boolean, required): Is a value or expression that can be evaluated to TRUE or FALSE

Input: `=NOT(TRUE)`

Observed output: `false`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>IFERROR — value returned</summary>

Returns value_if_error if expression is an error and the value of the expression itself otherwise

- value (value|range, required): Is any value or expression or reference
- valueIfError (value|range, required): Is any value or expression or reference

Input: `=IFERROR(A1:A5, A1:A5)`

Observed output: `1`

Classification: value returned. Elapsed: 7 ms.

</details>

<details>
<summary>XOR — value returned</summary>

Returns a logical 'Exclusive Or' of all arguments

- logical1 (boolean, required): Are 1 to 254 conditions you want to test that can be either TRUE or FALSE and can be logical values, arrays, or references
- logical2 (boolean, optional): TRUE or FALSE.
- rest (value|range, required): Value, reference, or range.

Input: `=XOR(TRUE, C1:C5, TRUE, C1:C5)`

Observed output: `false`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>IFNA — value returned</summary>

Returns the value you specify if the expression resolves to #N/A, otherwise returns the result of the expression

- value (value|range, required): Is any value or expression or reference
- valueIfNa (value|range, required): Is any value or expression or reference

Input: `=IFNA(A1:A5, A1:A5)`

Observed output: `1`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>IFS — value returned</summary>

Checks whether one or more conditions are met and returns a value corresponding to the first TRUE condition

- logicalTest (boolean, required): Is any value or expression that can be evaluated to TRUE or FALSE
- valueIfTrue (value|range, required): Is the value returned if Logical_test is TRUE
- rest (value|range, required): Value, reference, or range.

Input: `=IFS(TRUE, A1:A5, C1:C5, C1:C5)`

Observed output: `1`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>SWITCH — Excel error</summary>

Evaluates an expression against a list of values and returns the result corresponding to the first matching value. If there is no match, an optional default value is returned

- expression (value|range, required): Is an expression to be evaluated
- value1 (value|range, required): Is a value to be compared with expression
- result1 (value|range, required): Is a result to be returned if the corresponding value matches expression
- defaultOrValue2 (value|range, optional): Value argument for the calculation.
- result2 (value|range, optional): Value, reference, or range.
- rest (value|range, required): Value, reference, or range.

Input: `=SWITCH(A1, A1:A5, A1, C1:C5)`

Observed output: `"#N/A"`

Classification: Excel error. Elapsed: 9 ms.

</details>

<details>
<summary>LET — Excel error</summary>

Assigns calculation results to names. Useful for storing intermediate calculations and values by defining names inside a formula. These names only apply within the scope of the LET function.

- name1 (value|range, required): The name, or a calculation which can make use of all names within the LET. Names must start with a letter, cannot be the output of a formula, or conflict with range syntax.
- nameValue1 (value|range, required): The value associated with the name.
- calculationOrName2 (value|range, required): Value, reference, or range.
- nameValue2 (value|range, optional): Value argument for the calculation.
- rest (value|range, required): Value, reference, or range.

Input: `=LET(A1, A1, A1, C1:C5)`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 11 ms.

</details>

<details>
<summary>LAMBDA — Excel error</summary>

Creates a function value, which can be called within formulas

- parameterOrCalculation (value|range, required): A parameter, or calculation, calculating the result of the function. Parameters cannot be calculated. The last argument to LAMBDA must always be a calculation. The calculation, which may use the parameters, will return a function that can then be called.
- rest (value|range, required): Value, reference, or range.

Input: `=LAMBDA(A1, C1:C5, C1:C5)`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 11 ms.

</details>

<details>
<summary>MAKEARRAY — unimplemented text</summary>

Returns a calculated array of a specified row and column size, by applying a LAMBDA function.

- rows (value|range, required): Is the number of rows in the array. Must be greater than zero.
- columns (value|range, required): Is the number of columns in the array. Must be greater than zero.
- functionParam (value|range, required): Is a LAMBDA that is called to create the array. The LAMBDA takes two parameters, row index and column index.

Input: `=MAKEARRAY(A1, A1, A1)`

Observed output: `"MAKEARRAY is not implemented. rows=1, columns=1, functionParam=1"`

Classification: unimplemented text. Elapsed: 11 ms.

</details>

<details>
<summary>MAP — unimplemented text</summary>

Returns an array formed by 'mapping' each value in the array(s) to a new value by applying a lambda to create a new value.

- array (array, required): Is an array to be mapped
- lambdaOrArray2 (array, required): Is a LAMBDA which must be the last argument and must have a parameter for each array passed, or another array to be mapped
- rest (value|range, required): Value, reference, or range.

Input: `=MAP({1,2;3,4}, {1,2;3,4}, C1:C5, C1:C5)`

Observed output: `"MAP is not implemented. array=1,2,3,4, lambdaOrArray2=1,2,3,4, rest.length=2"`

Classification: unimplemented text. Elapsed: 14 ms.

</details>

<details>
<summary>REDUCE — unimplemented text</summary>

Reduces an array to an accumulated value by applying a LAMBDA function to each value and returning the total value in the accumulator.

- initialValue (value|range, required): Is the starting value for the accumulator
- array (array, required): Is an array to be reduced
- functionParam (value|range, required): Is a LAMBDA that is called to reduce the array. The LAMBDA takes two parameters, accumulator and value.

Input: `=REDUCE(A1, {1,2;3,4}, A1)`

Observed output: `"REDUCE is not implemented. initialValue=1, array=1,2,3,4, functionParam=1"`

Classification: unimplemented text. Elapsed: 15 ms.

</details>

<details>
<summary>SCAN — unimplemented text</summary>

Scans an array by applying a LAMBDA function to each value and returns an array that has each intermediate value

- initialValue (value|range, required): Is the the starting value for the accumulator
- array (array, required): Is an array to be scanned
- functionParam (value|range, required): Is a LAMBDA that is called to scan the array. The LAMBDA takes two parameters, accumulator and value.

Input: `=SCAN(A1, {1,2;3,4}, A1)`

Observed output: `"SCAN is not implemented. initialValue=1, array=1,2,3,4, functionParam=1"`

Classification: unimplemented text. Elapsed: 11 ms.

</details>

<details>
<summary>BYROW — Excel error</summary>

Applies a LAMBDA function to each row and returns an array of the results.

- array (array, required): Is an array to be separated by row
- functionParam (value|range, optional): Is a LAMBDA that is called to scan the array. The LAMBDA takes two parameters, accumulator and value.

Input: `=BYROW({1,2;3,4}, A1)`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 17 ms.

</details>

<details>
<summary>BYCOL — Excel error</summary>

Applies a LAMBDA function to each column and returns an array of the results.

- array (array, required): Is an array to be separated by column
- functionParam (value|range, optional): Is a LAMBDA that is called to scan the array. The LAMBDA takes two parameters, accumulator and value.

Input: `=BYCOL({1,2;3,4}, A1)`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 9 ms.

</details>


### Formula category: lookup-reference

<details>
<summary>ROW — value returned</summary>

Returns the row number of a reference

- reference (range, optional): Is the cell or a single range of cells for which you want the row number; if omitted, returns the cell containing the ROW function

Input: `=ROW(A1:A5)`

Observed output: `1`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>COLUMN — value returned</summary>

Returns the column number of a reference

- reference (range, optional): Is the cell or range of contiguous cells for which you want the column number. If omitted, the cell containing the COLUMN function is used

Input: `=COLUMN(A1:A5)`

Observed output: `1`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>LOOKUP — Excel error</summary>

Looks up a value either from a one-row or one-column range or from an array. Provided for backward compatibility

- lookupValue (value|range, required): Is a value that LOOKUP searches for in Lookup_vector and can be a number, text, a logical value, or a name or reference to a value
- lookupVector (value|range, required): Is a range that contains only one row or one column of text, numbers, or logical values, placed in ascending order
- resultVectorLookupValue (value|range, required): Is a range that contains only one row or column, the same size as Lookup_vector
- array (array, required): Is a value that LOOKUP searches for in Array and can be a number, text, a logical value, or a name or reference to a value is a range of cells that contain text, number, or logical values that you want to compare with Lookup_value

Input: `=LOOKUP(A2:A10, A2:A10, A2:A10, {1,2;3,4})`

Observed output: `"#N/A"`

Classification: Excel error. Elapsed: 10 ms.

</details>

<details>
<summary>INDEX — Excel error</summary>

Returns a value or reference of the cell at the intersection of a particular row and column, in a given range

- array (array, required): Is a range of cells or an array constant.
- rowNum (value|range, required): Selects the row in Array or Reference from which to return a value. If omitted, Column_num is required
- columnNumReference (range, required): Selects the column in Array or Reference from which to return a value. If omitted, Row_num is required
- rowNum2 (value|range, required): Is a reference to one or more cell ranges
- columnNum (value|range, optional): Selects the row in Array or Reference from which to return a value. If omitted, Column_num is required
- areaNum (value|range, optional): Selects the column in Array or Reference from which to return a value. If omitted, Row_num is required selects a range in Reference from which to return a value. The first area selected or entered is area 1, the second area is area 2, and so on

Input: `=INDEX({1,2;3,4}, A1, A1:A5, A1)`

Observed output: `"#REF!"`

Classification: Excel error. Elapsed: 9 ms.

</details>

<details>
<summary>MATCH — value returned</summary>

Returns the relative position of an item in an array that matches a specified value in a specified order

- lookupValue (value|range, required): Is the value you use to find the value you want in the array, a number, text, or logical value, or a reference to one of these
- lookupArray (array, required): Is a contiguous range of cells containing possible lookup values, an array of values, or a reference to an array
- matchType (value|range, optional): Is a number 1, 0, or -1 indicating which value to return.

Input: `=MATCH(A2:A10, {1,2;3,4}, A1)`

Observed output: `2`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>AREAS — unimplemented text</summary>

Returns the number of areas in a reference. An area is a range of contiguous cells or a single cell

- reference (range, required): Is a reference to a cell or range of cells and can refer to multiple areas

Input: `=AREAS(A1:A5)`

Observed output: `"AREAS is not implemented. reference=1,2,3,4,5"`

Classification: unimplemented text. Elapsed: 9 ms.

</details>

<details>
<summary>ROWS — value returned</summary>

Returns the number of rows in a reference or array

- array (array, required): Is an array, an array formula, or a reference to a range of cells for which you want the number of rows

Input: `=ROWS({1,2;3,4})`

Observed output: `2`

Classification: value returned. Elapsed: 6 ms.

</details>

<details>
<summary>COLUMNS — value returned</summary>

Returns the number of columns in an array or reference

- array (array, required): Is an array or array formula, or a reference to a range of cells for which you want the number of columns

Input: `=COLUMNS({1,2;3,4})`

Observed output: `2`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>OFFSET — unimplemented text</summary>

Returns a reference to a range that is a given number of rows and columns from a given reference

- reference (range, required): Is the reference from which you want to base the offset, a reference to a cell or range of adjacent cells
- rows (value|range, required): Is the number of rows, up or down, that you want the upper-left cell of the result to refer to
- cols (value|range, required): Is the number of columns, to the left or right, that you want the upper-left cell of the result to refer to
- height (value|range, optional): Is the height, in number of rows, that you want the result to be, the same height as Reference if omitted
- width (value|range, optional): Is the width, in number of columns, that you want the result to be, the same width as Reference if omitted

Input: `=OFFSET(A1:A5, A1, A1)`

Observed output: `"OFFSET is not implemented. reference=1,2,3,4,5, rows=1, cols=1, height=undefined, width=undefined"`

Classification: unimplemented text. Elapsed: 8 ms.

</details>

<details>
<summary>TRANSPOSE — value returned</summary>

Converts a vertical range of cells to a horizontal range, or vice versa

- array (array, required): Is a range of cells on a worksheet or an array of values that you want to transpose

Input: `=TRANSPOSE({1,2;3,4})`

Observed output: `1`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>CHOOSE — value returned</summary>

Chooses a value or action to perform from a list of values, based on an index number

- indexNum (value|range, required): Specifies which value argument is selected. Index_num must be between 1 and 254, or a formula or a reference to a number between 1 and 254
- value1 (value|range, required): Are 1 to 254 numbers, cell references, defined names, formulas, functions, or text arguments from which CHOOSE selects
- value2 (value|range, optional): Value argument for the calculation.
- rest (value|range, required): Value, reference, or range.

Input: `=CHOOSE(A1, A1:A5, C1:C5, C1:C5)`

Observed output: `1`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>HLOOKUP — value returned</summary>

Looks for a value in the top row of a table or array of values and returns the value in the same column from a row you specify

- lookupValue (value|range, required): Is the value to be found in the first row of the table and can be a value, a reference, or a text string
- tableArray (array, required): Is a table of text, numbers, or logical values in which data is looked up. Table_array can be a reference to a range or a range name
- rowIndexNum (value|range, required): Is the row number in table_array from which the matching value should be returned. The first row of values in the table is row 1
- rangeLookup (range, optional): Is a logical value: to find the closest match in the top row (sorted in ascending order) = TRUE or omitted; find an exact match = FALSE

Input: `=HLOOKUP(A2:A10, {1,2;3,4}, A1)`

Observed output: `2`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>VLOOKUP — value returned</summary>

Looks for a value in the leftmost column of a table, and then returns a value in the same row from a column you specify. By default, the table must be sorted in an ascending order

- lookupValue (value|range, required): Is the value to be found in the first column of the table, and can be a value, a reference, or a text string
- tableArray (array, required): Is a table of text, numbers, or logical values, in which data is retrieved. Table_array can be a reference to a range or a range name
- colIndexNum (value|range, required): Is the column number in table_array from which the matching value should be returned. The first column of values in the table is column 1
- rangeLookup (range, optional): Is a logical value: to find the closest match in the first column (sorted in ascending order) = TRUE or omitted; find an exact match = FALSE

Input: `=VLOOKUP(A2:A10, {1,2;3,4}, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>INDIRECT — unimplemented text</summary>

Returns the reference specified by a text string

- refText (string, required): Is a reference to a cell that contains an A1- or R1C1-style reference, a name defined as a reference, or a reference to a cell as a text string
- a1 (value|range, optional): Is a logical value that specifies the type of reference in Ref_text: R1C1-style = FALSE; A1-style = TRUE or omitted

Input: `=INDIRECT("text", A1)`

Observed output: `"INDIRECT is not implemented. refText=text, a1=1"`

Classification: unimplemented text. Elapsed: 13 ms.

</details>

<details>
<summary>ADDRESS — value returned</summary>

Creates a cell reference as text, given specified row and column numbers

- rowNum (value|range, required): Is the row number to use in the cell reference: Row_number = 1 for row 1
- columnNum (value|range, required): Is the column number to use in the cell reference. For example, Column_number = 4 for column D
- absNum (value|range, optional): Specifies the reference type: absolute = 1; absolute row/relative column = 2; relative row/absolute column = 3; relative = 4
- a1 (value|range, optional): Is a logical value that specifies the reference style: A1 style = 1 or TRUE; R1C1 style = 0 or FALSE
- sheetText (string, optional): Is text specifying the name of the worksheet to be used as the external reference

Input: `=ADDRESS(A1, A1, A1)`

Observed output: `"$A$1"`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>GETPIVOTDATA — unimplemented text</summary>

Extracts data stored in a PivotTable.

- dataField (value|range, required): Is the name of the data field to extract data from
- pivotTable (value|range, required): Is a reference to a cell or range of cells in the PivotTable that contains the data you want to retrieve
- field (value|range, optional): Field to refer to
- item (value|range, optional): Field item to refer to
- rest (value|range, required): Value, reference, or range.

Input: `=GETPIVOTDATA(A1, A1, C1:C5, C1:C5)`

Observed output: `"GETPIVOTDATA is not implemented. dataField=1, pivotTable=1, field=3,6,9,12,15, item=3,6,9,12,15, rest.length=0"`

Classification: unimplemented text. Elapsed: 9 ms.

</details>

<details>
<summary>HYPERLINK — unimplemented text</summary>

Creates a shortcut or jump that opens a document stored on your hard drive, a network server, or on the Internet

- linkLocation (value|range, required): Is the text giving the path and file name to the document to be opened, a hard drive location, UNC address, or URL path
- friendlyName (value|range, optional): Is text or a number that is displayed in the cell. If omitted, the cell displays the Link_location text

Input: `=HYPERLINK(A1, A1)`

Observed output: `"HYPERLINK is not implemented. linkLocation=1, friendlyName=1"`

Classification: unimplemented text. Elapsed: 11 ms.

</details>

<details>
<summary>FORMULATEXT — unimplemented text</summary>

Returns a formula as a string

- reference (range, required): Is a reference to a formula

Input: `=FORMULATEXT(A1:A5)`

Observed output: `"FORMULATEXT is not implemented. reference=1,2,3,4,5"`

Classification: unimplemented text. Elapsed: 11 ms.

</details>

<details>
<summary>FIELDVALUE — unimplemented text</summary>

Extracts a value from a field of a given record

- value (value|range, required): The record from which you want to extract the field
- fieldName (value|range, required): The names of the field or fields that you want to extract

Input: `=FIELDVALUE(A1:A5, A1)`

Observed output: `"FIELDVALUE is not implemented. value=1,2,3,4,5, fieldName=1"`

Classification: unimplemented text. Elapsed: 7 ms.

</details>

<details>
<summary>FILTER — value returned</summary>

Filter a range or array

- array (array, required): The range or array to filter
- include (value|range, required): An array of booleans where TRUE represents a row or column to retain
- ifEmpty (value|range, optional): Returned if no items are retained

Input: `=FILTER({1,2;3,4}, A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>SORT — value returned</summary>

Sorts a range or array

- array (array, required): The range or array to sort
- sortIndex (value|range, optional): A number indicating the row or column to sort by
- sortOrder (value|range, optional): A number indicating the desired sort order; 1 for ascending order (default), -1 for descending order
- byCol (value|range, optional): A logical value indicating the desired sort direction: FALSE to sort by row (default), TRUE to sort by column

Input: `=SORT({1,2;3,4}, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>UNIQUE — value returned</summary>

Returns the unique values from a range or array.

- array (array, required): The range or array from which to return unique rows or columns
- byCol (value|range, optional): Is a logical value: compare rows against each other and return the unique rows = FALSE or omitted; compare columns against each other and return the unique columns = TRUE
- exactlyOnce (value|range, optional): Is a logical value: return rows or columns that occur exactly once from the array = TRUE; return all distinct rows or columns from the array = FALSE or omitted

Input: `=UNIQUE({1,2;3,4}, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 15 ms.

</details>

<details>
<summary>XMATCH — Excel error</summary>

Returns the relative position of an item in an array. By default, an exact match is required

- lookupValue (value|range, required): Is the value to search for
- lookupArray (array, required): Is the array or range to search
- matchMode (value|range, optional): Specify how to match the lookup_value against the values in lookup_array
- searchMode (value|range, optional): Specify the search mode to use. By default, a first to last search will be used

Input: `=XMATCH(A2:A10, {1,2;3,4}, A1)`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 12 ms.

</details>

<details>
<summary>XLOOKUP — Excel error</summary>

Searches a range or an array for a match and returns the corresponding item from a second range or array. By default, an exact match is used

- lookupValue (value|range, required): Is the value to search for
- lookupArray (array, required): Is the array or range to search
- returnArray (array, required): Is the array or range to return
- ifNotFound (value|range, optional): Returned if no match is found
- matchMode (value|range, optional): Specify how to match lookup_value against the values in lookup_array
- searchMode (value|range, optional): Specify the search mode to use. By default, a first to last search will be used

Input: `=XLOOKUP(A2:A10, {1,2;3,4}, {1,2;3,4})`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 10 ms.

</details>

<details>
<summary>SORTBY — Excel error</summary>

Sorts a range or array based on the values in a corresponding range or array

- array (array, required): The range or array to sort
- byArray (array, required): The range or array to sort on
- sortOrder (value|range, optional): A number indicating the desired sort order; 1 for ascending order (default), -1 for descending order
- rest (value|range, required): Value, reference, or range.

Input: `=SORTBY({1,2;3,4}, {1,2;3,4}, C1:C5, C1:C5)`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 7 ms.

</details>

<details>
<summary>WRAPROWS — value returned</summary>

Wraps a row or column vector after a specified number of values.

- vector (value|range, required): The vector or reference to wrap.
- wrapCount (value|range, required): The maximum number of values per row.
- padWith (value|range, optional): The value with which to pad. The default is #N/A.

Input: `=WRAPROWS(A1, 10, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>VSTACK — value returned</summary>

Vertically stacks arrays into one array.

- array1 (array, required): An array or reference to be stacked.
- array2 (array, optional): Array or spill range input.
- rest (value|range, required): Value, reference, or range.

Input: `=VSTACK({1,2;3,4}, C1:C5, {1,2;3,4}, C1:C5)`

Observed output: `1`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>HSTACK — value returned</summary>

Horizontally stacks arrays into one array.

- array1 (array, required): An array or reference to be stacked.
- array2 (array, optional): Array or spill range input.
- rest (value|range, required): Value, reference, or range.

Input: `=HSTACK({1,2;3,4}, C1:C5, {1,2;3,4}, C1:C5)`

Observed output: `1`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>CHOOSEROWS — Excel error</summary>

Returns rows from an array or reference.

- array (array, required): The array or reference containing the rows to be returned.
- rowNum1 (value|range, required): The number of the row to be returned.
- rowNum2 (value|range, optional): Value, reference, or range.
- rest (value|range, required): Value, reference, or range.

Input: `=CHOOSEROWS({1,2;3,4}, A1, C1:C5, C1:C5)`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 12 ms.

</details>

<details>
<summary>CHOOSECOLS — Excel error</summary>

Returns columns from an array or reference.

- array (array, required): The array or reference containing the columns to be returned.
- colNum1 (value|range, required): The number of the column to be returned.
- colNum2 (value|range, optional): Value, reference, or range.
- rest (value|range, required): Value, reference, or range.

Input: `=CHOOSECOLS({1,2;3,4}, A1, C1:C5, C1:C5)`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 14 ms.

</details>

<details>
<summary>TOCOL — value returned</summary>

Returns the array as one column.

- array (array, required): The array or reference to return as a column.
- ignore (value|range, optional): Whether to ignore certain types of values. By default, no values are ignored.
- scanByColumn (value|range, optional): Scan the array by column. By default, the array is scanned by row.

Input: `=TOCOL({1,2;3,4}, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>TOROW — value returned</summary>

Returns the array as one row.

- array (array, required): The array or reference to return as a row.
- ignore (value|range, optional): Whether to ignore certain types of values. By default, no values are ignored.
- scanByColumn (value|range, optional): Scan the array by column. By default, the array is scanned by row.

Input: `=TOROW({1,2;3,4}, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>WRAPCOLS — value returned</summary>

Wraps a row or column vector after a specified number of values.

- vector (value|range, required): The vector or reference to wrap.
- wrapCount (value|range, required): The maximum number of values per column.
- padWith (value|range, optional): The value with which to pad. The default is #N/A.

Input: `=WRAPCOLS(A1, 10, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>EXPAND — Excel error</summary>

Expands an array to the specified dimensions.

- array (array, required): The array to expand.
- rows (value|range, required): The number of rows in the expanded array. If missing, rows will not be expanded.
- columns (value|range, optional): The number of columns in the expanded array. If missing, columns will not be expanded.
- padWith (value|range, optional): The value with which to pad. The default is #N/A.

Input: `=EXPAND({1,2;3,4}, A1, A1)`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 9 ms.

</details>

<details>
<summary>TAKE — value returned</summary>

Returns rows or columns from array start or end.

- array (array, required): The array from which to take rows or columns.
- rows (value|range, required): The number of rows to take. A negative value takes from the end of the array.
- columns (value|range, optional): The number of columns to take. A negative value takes from the end of the array.

Input: `=TAKE({1,2;3,4}, A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>DROP — value returned</summary>

Drops rows or columns from array start or end.

- array (array, required): The array from which to drop rows or columns.
- rows (value|range, optional): The number of rows to drop. A negative value drops from the end of the array.
- columns (value|range, optional): The number of columns to drop. A negative value drops from the end of the array.

Input: `=DROP({1,2;3,4}, A1)`

Observed output: `3`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>PY — unimplemented text</summary>

Executes Python code that returns a value into the cell (not yet supported in Granola).

- useCtrlEnterToCommitPythonCodeCreatesPythonFormulas (value|range, required): Value, reference, or range.

Input: `=PY(A1)`

Observed output: `"PY is not implemented. useCtrlEnterToCommitPythonCodeCreatesPythonFormulas=1"`

Classification: unimplemented text. Elapsed: 17 ms.

</details>

<details>
<summary>IMAGE — unimplemented text</summary>

Returns an image from a given source

- source (value|range, required): The path of the source that points to the image
- altText (string, optional): The alternative text that describes the image for accessibility
- sizing (value|range, optional): The setting that determines the dimensions in which the image will be rendered in the cell
- height (value|range, optional): The custom height of the image in pixels
- width (value|range, optional): The custom width of the image in pixels

Input: `=IMAGE(A1, "text")`

Observed output: `"IMAGE is not implemented. source=1, altText=text, sizing=undefined, height=undefined, width=undefined"`

Classification: unimplemented text. Elapsed: 14 ms.

</details>

<details>
<summary>GROUPBY — unimplemented text</summary>

Aggregate values by row fields

- rowFields (value|range, required): The range or array containing the row fields
- values (value|range, required): The range or array containing the values to aggregate
- functionParam (value|range, required): The function used to do the aggregations
- fieldHeaders (value|range, optional): A number between 0 and 3 that specifies whether the field data has headers and whether field headers should be returned in the results
- totalDepth (value|range, optional): Show totals for the row fields. 0 for none, 1 for grand totals, 2 for grand total and first level subtotals, etc
- sortOrder (value|range, optional): The column index to sort the row fields on. A negative index will sort in reverse order
- filterArray (array, optional): An array of booleans where TRUE represents a row to retain
- fieldRelationship (value|range, optional): The relationship of fields when multiple columns are supplied to the row_fields argument

Input: `=GROUPBY(A1, A1:A5, A1)`

Observed output: `"GROUPBY is not implemented. rowFields=1, values=1,2,3,4,5, functionParam=1, fieldHeaders=undefined, totalDepth=undefined, sortOrder=undefined, filterArray=undefined, fieldRelationship=undefined"`

Classification: unimplemented text. Elapsed: 7 ms.

</details>

<details>
<summary>PIVOTBY — unimplemented text</summary>

Aggregate values by rows and columns

- rowFields (value|range, required): The range or array containing the row fields
- colFields (value|range, required): The range or array containing the column fields
- values (value|range, required): The range or array containing the values to aggregate
- functionParam (value|range, required): The function used to do the aggregations
- fieldHeaders (value|range, optional): A number between 0 and 3 that specifies whether the field data has headers and whether field headers should be returned in the results
- rowTotalDepth (value|range, optional): Show totals for the row fields. 0 for none, 1 for grand totals, 2 for grand total and first level subtotals, etc
- rowSortOrder (value|range, optional): The column index to sort the row fields on. A negative index will sort in reverse order
- colTotalDepth (value|range, optional): Show totals for the column fields. 0 for none, 1 for grand totals, 2 for grand total and first level subtotals, etc
- colSortOrder (value|range, optional): The column index to sort the column fields on. A negative index will sort in reverse order
- filterArray (array, optional): An array of booleans where TRUE represents a row to retain
- relativeTo (value|range, optional): Defines what PERCENTOF is calculated relative to. This

Input: `=PIVOTBY(A1, A1, A1:A5, A1)`

Observed output: `"PIVOTBY is not implemented. rowFields=1, colFields=1, values=1,2,3,4,5, functionParam=1, fieldHeaders=undefined, rowTotalDepth=undefined, rowSortOrder=undefined, colTotalDepth=undefined, colSortOrder=undefined, filterArray=undefined, relativeTo=undefined"`

Classification: unimplemented text. Elapsed: 11 ms.

</details>

<details>
<summary>TRIMRANGE — unimplemented text</summary>

Trims a range to the last used cell in any direction.

- range (range, required): The range to be trimmed
- rowTrimMode (value|range, optional): Row trim direction
- colTrimMode (value|range, optional): Column trim direction

Input: `=TRIMRANGE(A1:A5, A1)`

Observed output: `"TRIMRANGE is not implemented. range=1,2,3,4,5, rowTrimMode=1, colTrimMode=undefined"`

Classification: unimplemented text. Elapsed: 7 ms.

</details>


### Formula category: math-trig

<details>
<summary>SUM — value returned</summary>

Adds all the numbers in a range of cells

- number1 (number|range, required): Are 1 to 255 numbers to sum. Logical values and text are ignored in cells, included if typed as arguments
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=SUM(10, C1:C5, 20, C1:C5)`

Observed output: `120`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>SIN — value returned</summary>

Returns the sine of an angle

- numberParam (number|range, required): Is the angle in radians for which you want the sine. Degrees * PI()/180 = radians

Input: `=SIN(10)`

Observed output: `-0.5440211108893698`

Classification: value returned. Elapsed: 7 ms.

</details>

<details>
<summary>COS — value returned</summary>

Returns the cosine of an angle

- numberParam (number|range, required): Is the angle in radians for which you want the cosine

Input: `=COS(10)`

Observed output: `-0.8390715290764524`

Classification: value returned. Elapsed: 7 ms.

</details>

<details>
<summary>TAN — value returned</summary>

Returns the tangent of an angle

- numberParam (number|range, required): Is the angle in radians for which you want the tangent. Degrees * PI()/180 = radians

Input: `=TAN(10)`

Observed output: `0.6483608274590866`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>ATAN — value returned</summary>

Returns the arctangent of a number in radians, in the range -Pi/2 to Pi/2

- numberParam (number|range, required): Is the tangent of the angle you want

Input: `=ATAN(10)`

Observed output: `1.4711276743037347`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>PI — value returned</summary>

Returns the value of Pi, 3.14159265358979, accurate to 15 digits


Input: `=PI()`

Observed output: `3.141592653589793`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>SQRT — value returned</summary>

Returns the square root of a number

- numberParam (number|range, required): Is the number for which you want the square root

Input: `=SQRT(10)`

Observed output: `3.1622776601683795`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>EXP — value returned</summary>

Returns e raised to the power of a given number

- numberParam (number|range, required): Is the exponent applied to the base e. The constant e equals 2.71828182845904, the base of the natural logarithm

Input: `=EXP(10)`

Observed output: `22026.465794806718`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>LN — value returned</summary>

Returns the natural logarithm of a number

- numberParam (number|range, required): Is the positive real number for which you want the natural logarithm

Input: `=LN(10)`

Observed output: `2.302585092994046`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>LOG10 — value returned</summary>

Returns the base-10 logarithm of a number

- numberParam (number|range, required): Is the positive real number for which you want the base-10 logarithm

Input: `=LOG10(10)`

Observed output: `1`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>ABS — value returned</summary>

Returns the absolute value of a number, a number without its sign

- numberParam (number|range, required): Is the real number for which you want the absolute value

Input: `=ABS(10)`

Observed output: `10`

Classification: value returned. Elapsed: 6 ms.

</details>

<details>
<summary>INT — value returned</summary>

Rounds a number down to the nearest integer

- numberParam (number|range, required): Is the real number you want to round down to an integer

Input: `=INT(10)`

Observed output: `10`

Classification: value returned. Elapsed: 16 ms.

</details>

<details>
<summary>SIGN — value returned</summary>

Returns the sign of a number: 1 if the number is positive, zero if the number is zero, or -1 if the number is negative

- numberParam (number|range, required): Is any real number

Input: `=SIGN(10)`

Observed output: `1`

Classification: value returned. Elapsed: 18 ms.

</details>

<details>
<summary>ROUND — value returned</summary>

Rounds a number to a specified number of digits

- numberParam (number|range, required): Is the number you want to round
- numDigits (value|range, required): Is the number of digits to which you want to round. Negative rounds to the left of the decimal point; zero to the nearest integer

Input: `=ROUND(10, A1)`

Observed output: `10`

Classification: value returned. Elapsed: 16 ms.

</details>

<details>
<summary>MOD — value returned</summary>

Returns the remainder after a number is divided by a divisor

- numberParam (number|range, required): Is the number for which you want to find the remainder after the division is performed
- divisor (value|range, required): Is the number by which you want to divide Number

Input: `=MOD(10, A1)`

Observed output: `0`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>RAND — value returned</summary>

Returns a random number greater than or equal to 0 and less than 1, evenly distributed (changes on recalculation)


Input: `=RAND()`

Observed output: `0.11945641925007944`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>ATAN2 — value returned</summary>

Returns the arctangent of the specified x- and y- coordinates, in radians between -Pi and Pi, excluding -Pi

- xNum (value|range, required): Is the x-coordinate of the point
- yNum (value|range, required): Is the y-coordinate of the point

Input: `=ATAN2(A1, A1)`

Observed output: `0.7853981633974483`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>ASIN — Excel error</summary>

Returns the arcsine of a number in radians, in the range -Pi/2 to Pi/2

- numberParam (number|range, required): Is the sine of the angle you want and must be from -1 to 1

Input: `=ASIN(10)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 11 ms.

</details>

<details>
<summary>ACOS — Excel error</summary>

Returns the arccosine of a number, in radians in the range 0 to Pi. The arccosine is the angle whose cosine is Number

- numberParam (number|range, required): Is the cosine of the angle you want and must be from -1 to 1

Input: `=ACOS(10)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 9 ms.

</details>

<details>
<summary>LOG — Excel error</summary>

Returns the logarithm of a number to the base you specify

- numberParam (number|range, required): Is the positive real number for which you want the logarithm
- base (value|range, optional): Is the base of the logarithm; 10 if omitted

Input: `=LOG(10, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 11 ms.

</details>

<details>
<summary>MDETERM — value returned</summary>

Returns the matrix determinant of an array

- array (array, required): Is a numeric array with an equal number of rows and columns, either a cell range or an array constant

Input: `=MDETERM({1,2;3,4})`

Observed output: `-2`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>MINVERSE — value returned</summary>

Returns the inverse matrix for the matrix stored in an array

- array (array, required): Is a numeric array with an equal number of rows and columns, either a cell range or an array constant

Input: `=MINVERSE({1,2;3,4})`

Observed output: `-2`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>MMULT — value returned</summary>

Returns the matrix product of two arrays, an array with the same number of rows as array1 and columns as array2

- array1 (array, required): Is the first array of numbers to multiply and must have the same number of columns as Array2 has rows
- array2 (array, required): Array or spill range input.

Input: `=MMULT({1,2;3,4}, {1,2;3,4})`

Observed output: `7`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>PRODUCT — value returned</summary>

Multiplies all the numbers given as arguments

- number1 (number|range, required): Are 1 to 255 numbers, logical values, or text representations of numbers that you want to multiply
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=PRODUCT(10, C1:C5, 20, C1:C5)`

Observed output: `170061120000`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>FACT — value returned</summary>

Returns the factorial of a number, equal to 1*2*3*...* Number

- numberParam (number|range, required): Is the nonnegative number you want the factorial of

Input: `=FACT(10)`

Observed output: `3628800`

Classification: value returned. Elapsed: 7 ms.

</details>

<details>
<summary>TRUNC — value returned</summary>

Truncates a number to an integer by removing the decimal, or fractional, part of the number

- numberParam (number|range, required): Is the number you want to truncate
- numDigits (value|range, optional): Is a number specifying the precision of the truncation, 0 (zero) if omitted

Input: `=TRUNC(10, A1)`

Observed output: `10`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>ROUNDUP — value returned</summary>

Rounds a number up, away from zero

- numberParam (number|range, required): Is any real number that you want rounded up
- numDigits (value|range, required): Is the number of digits to which you want to round. Negative rounds to the left of the decimal point; zero or omitted, to the nearest integer

Input: `=ROUNDUP(10, A1)`

Observed output: `10`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>ROUNDDOWN — value returned</summary>

Rounds a number down, toward zero

- numberParam (number|range, required): Is any real number that you want rounded down
- numDigits (value|range, required): Is the number of digits to which you want to round. Negative rounds to the left of the decimal point; zero or omitted, to the nearest integer

Input: `=ROUNDDOWN(10, A1)`

Observed output: `10`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>SUMPRODUCT — Excel error</summary>

Returns the sum of the products of corresponding ranges or arrays

- array1 (array, required): Are 2 to 255 arrays for which you want to multiply and then add components. All arrays must have the same dimensions
- rest (value|range, required): Value, reference, or range.

Input: `=SUMPRODUCT({1,2;3,4}, C1:C5, C1:C5)`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 13 ms.

</details>

<details>
<summary>SINH — value returned</summary>

Returns the hyperbolic sine of a number

- numberParam (number|range, required): Is any real number

Input: `=SINH(10)`

Observed output: `11013.232874703393`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>COSH — value returned</summary>

Returns the hyperbolic cosine of a number

- numberParam (number|range, required): Is any real number

Input: `=COSH(10)`

Observed output: `11013.232920103324`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>TANH — value returned</summary>

Returns the hyperbolic tangent of a number

- numberParam (number|range, required): Is any real number

Input: `=TANH(10)`

Observed output: `0.9999999958776927`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>ASINH — value returned</summary>

Returns the inverse hyperbolic sine of a number

- numberParam (number|range, required): Is any real number equal to or greater than 1

Input: `=ASINH(10)`

Observed output: `2.99822295029797`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>ACOSH — value returned</summary>

Returns the inverse hyperbolic cosine of a number

- numberParam (number|range, required): Is any real number equal to or greater than 1

Input: `=ACOSH(10)`

Observed output: `2.993222846126381`

Classification: value returned. Elapsed: 7 ms.

</details>

<details>
<summary>ATANH — Excel error</summary>

Returns the inverse hyperbolic tangent of a number

- numberParam (number|range, required): Is any real number between -1 and 1 excluding -1 and 1

Input: `=ATANH(10)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 10 ms.

</details>

<details>
<summary>COMBIN — value returned</summary>

Returns the number of combinations for a given number of items

- numberParam (number|range, required): Is the total number of items
- numberChosen (number|range, required): Is the number of items in each combination

Input: `=COMBIN(10, 10)`

Observed output: `1`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>EVEN — value returned</summary>

Rounds a positive number up and negative number down to the nearest even integer

- numberParam (number|range, required): Is the value to round

Input: `=EVEN(10)`

Observed output: `10`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>ODD — value returned</summary>

Rounds a positive number up and negative number down to the nearest odd integer

- numberParam (number|range, required): Is the value to round

Input: `=ODD(10)`

Observed output: `11`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>SUMXMY2 — value returned</summary>

Sums the squares of the differences in two corresponding ranges or arrays

- arrayX (array, required): Is the first range or array of values and can be a number or name, array, or reference that contains numbers
- arrayY (array, required): Is the second range or array of values and can be a number or name, array, or reference that contains numbers

Input: `=SUMXMY2({1,2;3,4}, {1,2;3,4})`

Observed output: `0`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>SUMX2MY2 — value returned</summary>

Sums the differences between the squares of two corresponding ranges or arrays

- arrayX (array, required): Is the first range or array of numbers and can be a number or name, array, or reference that contains numbers
- arrayY (array, required): Is the second range or array of numbers and can be a number or name, array, or reference that contains numbers

Input: `=SUMX2MY2({1,2;3,4}, {1,2;3,4})`

Observed output: `0`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>SUMX2PY2 — value returned</summary>

Returns the sum total of the sums of squares of numbers in two corresponding ranges or arrays

- arrayX (array, required): Is the first range or array of numbers and can be a number or name, array, or reference that contains numbers
- arrayY (array, required): Is the second range or array of numbers and can be a number or name, array, or reference that contains numbers

Input: `=SUMX2PY2({1,2;3,4}, {1,2;3,4})`

Observed output: `60`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>SUMSQ — value returned</summary>

Returns the sum of the squares of the arguments. The arguments can be numbers, arrays, names, or references to cells that contain numbers

- number1 (number|range, required): Are 1 to 255 numbers, arrays, names, or references to arrays for which you want the sum of the squares
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=SUMSQ(10, C1:C5, 20, C1:C5)`

Observed output: `1490`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>POWER — value returned</summary>

Returns the result of a number raised to a power

- numberParam (number|range, required): Is the base number, any real number
- power (value|range, required): Is the exponent, to which the base number is raised

Input: `=POWER(10, A1)`

Observed output: `10`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>RADIANS — value returned</summary>

Converts degrees to radians

- angle (value|range, required): Is an angle in degrees that you want to convert

Input: `=RADIANS(A1)`

Observed output: `0.017453292519943295`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>DEGREES — value returned</summary>

Converts radians to degrees

- angle (value|range, required): Is the angle in radians that you want to convert

Input: `=DEGREES(A1)`

Observed output: `57.29577951308232`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>SUBTOTAL — value returned</summary>

Returns a subtotal in a list or database

- functionNum (value|range, required): Is the number 1 to 11 that specifies the summary function for the subtotal.
- ref1 (value|range, required): Are 1 to 254 ranges or references for which you want the subtotal
- rest (value|range, required): Value, reference, or range.

Input: `=SUBTOTAL(A1, A1, C1:C5, C1:C5)`

Observed output: `1`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>SUMIF — value returned</summary>

Adds the cells specified by a given condition or criteria

- range (range, required): Is the range of cells you want evaluated
- criteria (criteria, required): Is the condition or criteria in the form of a number, expression, or text that defines which cells will be added
- sumRange (range, optional): Are the actual cells to sum. If omitted, the cells in range are used

Input: `=SUMIF(A1:A5, ">0", A1:A5)`

Observed output: `15`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>ROMAN — value returned</summary>

Converts an Arabic numeral to Roman, as text

- numberParam (number|range, required): Is the Arabic numeral you want to convert
- form (value|range, optional): Is the number specifying the type of Roman numeral you want.

Input: `=ROMAN(10, A1)`

Observed output: `"X"`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>SERIESSUM — Excel error</summary>

Returns the sum of a power series based on the formula

- x (value|range, required): Is the input value to the power series
- n (value|range, required): Is the initial power to which you want to raise x
- m (value|range, required): Is the step by which to increase n for each term in the series
- coefficients (value|range, required): Is a set of coefficients by which each successive power of x is multiplied

Input: `=SERIESSUM(A1, A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 11 ms.

</details>

<details>
<summary>FACTDOUBLE — value returned</summary>

Returns the double factorial of a number

- numberParam (number|range, required): Is the value for which to return the double factorial

Input: `=FACTDOUBLE(10)`

Observed output: `3840`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>SQRTPI — value returned</summary>

Returns the square root of (number * Pi)

- numberParam (number|range, required): Is the number by which p is multiplied

Input: `=SQRTPI(10)`

Observed output: `5.604991216397929`

Classification: value returned. Elapsed: 7 ms.

</details>

<details>
<summary>QUOTIENT — value returned</summary>

Returns the integer portion of a division

- numerator (value|range, required): Is the dividend
- denominator (value|range, required): Is the divisor

Input: `=QUOTIENT(A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>MROUND — value returned</summary>

Returns a number rounded to the desired multiple

- numberParam (number|range, required): Is the value to round
- multiple (value|range, required): Is the multiple to which you want to round number

Input: `=MROUND(10, A1)`

Observed output: `10`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>RANDBETWEEN — value returned</summary>

Returns a random number between the numbers you specify

- bottom (value|range, required): Is the smallest integer RANDBETWEEN will return
- top (value|range, required): Is the largest integer RANDBETWEEN will return

Input: `=RANDBETWEEN(A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>GCD — value returned</summary>

Returns the greatest common divisor

- number1 (number|range, required): Are 1 to 255 values
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=GCD(10, C1:C5, 20, C1:C5)`

Observed output: `1`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>MULTINOMIAL — value returned</summary>

Returns the multinomial of a set of numbers

- number1 (number|range, required): Are 1 to 255 values for which you want the multinomial
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=MULTINOMIAL(10, C1:C5, 20, C1:C5)`

Observed output: `7.858496926878887e+113`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>LCM — value returned</summary>

Returns the least common multiple

- number1 (number|range, required): Are 1 to 255 values for which you want the least common multiple
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=LCM(10, C1:C5, 20, C1:C5)`

Observed output: `180`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>SUMIFS — value returned</summary>

Adds the cells specified by a given set of conditions or criteria

- sumRange (range, required): Are the actual cells to sum.
- criteriaRange (range, required): Is the range of cells you want evaluated for the particular condition
- criteria (criteria, required): Is the condition or criteria in the form of a number, expression, or text that defines which cells will be added
- rest (value|range, required): Value, reference, or range.

Input: `=SUMIFS(A1:A5, A1:A5, ">0", C1:C5)`

Observed output: `15`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>AGGREGATE — value returned</summary>



- functionNum (value|range, required): Value, reference, or range.
- options (value|range, required): Value, reference, or range.
- ref1 (value|range, required): Value, reference, or range.
- rest (value|range, required): Value, reference, or range.

Input: `=AGGREGATE(A1, A1, A1, C1:C5)`

Observed output: `1`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>ACOT — value returned</summary>

Returns the arccotangent of a number, in radians in the range 0 to Pi.

- numberParam (number|range, required): Is the cotangent of the angle you want

Input: `=ACOT(10)`

Observed output: `0.09966865249116204`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>ACOTH — value returned</summary>

Returns the inverse hyperbolic cotangent of a number

- numberParam (number|range, required): Is the hyperbolic cotangent of the angle that you want

Input: `=ACOTH(10)`

Observed output: `0.10033534773107562`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>COT — value returned</summary>

Returns the cotangent of an angle

- numberParam (number|range, required): Is the angle in radians for which you want the cotangent

Input: `=COT(10)`

Observed output: `1.5423510453569202`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>COTH — value returned</summary>

Returns the hyperbolic cotangent of a number

- numberParam (number|range, required): Is the angle in radians for which you want the hyperbolic cotangent

Input: `=COTH(10)`

Observed output: `1.0000000041223072`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>CSC — value returned</summary>

Returns the cosecant of an angle

- numberParam (number|range, required): Is the angle in radians for which you want the cosecant

Input: `=CSC(10)`

Observed output: `-1.8381639608896658`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>CSCH — value returned</summary>

Returns the hyperbolic cosecant of an angle

- numberParam (number|range, required): Is the angle in radians for which you want the hyperbolic cosecant

Input: `=CSCH(10)`

Observed output: `0.00009079985971212217`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>SEC — value returned</summary>

Returns the secant of an angle

- numberParam (number|range, required): Is the angle in radians for which you want the secant

Input: `=SEC(10)`

Observed output: `-1.1917935066878957`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>SECH — value returned</summary>

Returns the hyperbolic secant of an angle

- numberParam (number|range, required): Is the angle in radians for which you want the hyperbolic secant

Input: `=SECH(10)`

Observed output: `0.00009079985933781724`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>COMBINA — value returned</summary>

Returns the number of combinations with repetitions for a given number of items

- numberParam (number|range, required): Is the total number of items
- numberChosen (number|range, required): Is the number of items in each combination

Input: `=COMBINA(10, 10)`

Observed output: `92378`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>BASE — Excel error</summary>

Converts a number into a text representation with the given radix (base)

- numberParam (number|range, required): Is the number that you want to convert
- radix (value|range, required): Is the base Radix that you want to convert the number into
- minLength (value|range, optional): Is the minimum length of the returned string. If omitted leading zeros are not added

Input: `=BASE(10, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 14 ms.

</details>

<details>
<summary>DECIMAL — Excel error</summary>

Converts a text representation of a number in a given base into a decimal number

- numberParam (number|range, required): Is the number that you want to convert
- radix (value|range, required): Is the base Radix of the number you are converting

Input: `=DECIMAL(10, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 14 ms.

</details>

<details>
<summary>MUNIT — value returned</summary>

Returns the unit matrix for the specified dimension

- dimension (value|range, required): Is an integer specifying the dimension of the unit matrix that you want to return

Input: `=MUNIT(A1)`

Observed output: `1`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>ARABIC — Excel error</summary>

Converts a Roman numeral to Arabic

- text (string, required): Is the Roman numeral you want to convert

Input: `=ARABIC("text")`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 12 ms.

</details>

<details>
<summary>CEILING.MATH — value returned</summary>

Rounds a number up, to the nearest integer or to the nearest multiple of significance

- numberParam (number|range, required): Is the value you want to round
- significance (value|range, optional): Is the multiple to which you want to round
- mode (value|range, optional): When given and nonzero this function will round away from zero

Input: `=CEILING.MATH(10, A1)`

Observed output: `10`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>FLOOR.MATH — value returned</summary>

Rounds a number down, to the nearest integer or to the nearest multiple of significance

- numberParam (number|range, required): Is the value you want to round
- significance (value|range, optional): Is the multiple to which you want to round
- mode (value|range, optional): When given and nonzero this function will round towards zero

Input: `=FLOOR.MATH(10, A1)`

Observed output: `10`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>RANDARRAY — value returned</summary>

Returns an array of random numbers

- rows (value|range, optional): The number of rows in the returned array
- columns (value|range, optional): The number of columns in the returned array
- min (value|range, optional): The minimum number you would like returned
- max (value|range, optional): The maximum number you would like returned
- integer (value|range, optional): Return an integer or a decimal value. TRUE for an integer, FALSE for a decimal number

Input: `=RANDARRAY(A1)`

Observed output: `0.345885927460773`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>SEQUENCE — value returned</summary>

Returns a sequence of numbers

- rows (value|range, required): The number of rows to return
- columns (value|range, optional): The number of columns to return
- start (value|range, optional): The first number in the sequence
- step (value|range, optional): The amount to increment each subsequent value in the sequence

Input: `=SEQUENCE(A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>PERCENTOF — unimplemented text</summary>

Returns the percentage of a subset of a given data set

- dataSubset (value|range, required): Is the subset
- dataAll (value|range, required): Is the data set

Input: `=PERCENTOF(A1, A1)`

Observed output: `"PERCENTOF is not implemented. dataSubset=1, dataAll=1"`

Classification: unimplemented text. Elapsed: 9 ms.

</details>


### Formula category: statistical

<details>
<summary>COUNT — value returned</summary>

Counts the number of cells in a range that contain numbers

- value1 (value|range, required): Are 1 to 255 arguments that can contain or refer to a variety of different types of data, but only numbers are counted
- value2 (value|range, optional): Value argument for the calculation.
- rest (value|range, required): Value, reference, or range.

Input: `=COUNT(A1:A5, C1:C5, B1:B5, C1:C5)`

Observed output: `20`

Classification: value returned. Elapsed: 18 ms.

</details>

<details>
<summary>AVERAGE — value returned</summary>

Returns the arithmetic mean of the supplied numbers, ignoring text and empty cells.

- number1 (number|range, required): First number, reference, or range that contains values to average.
- number2 (number|range, optional): Additional number, reference, or range to include in the mean.
- rest (number|range, required): Optional extra arguments (up to 255 total) supplied as numbers or ranges.

Input: `=AVERAGE(A2:A6, 10, 12)`

Observed output: `6`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>MIN — value returned</summary>

Returns the smallest numeric value from the supplied numbers or ranges, ignoring text, logical values, and blanks.

- number1 (number|range, required): First number, reference, or range that contains values to compare.
- number2 (number|range, optional): Optional additional number, reference, or range.
- rest (number|range, required): Optional extra numbers or ranges (up to 255 total arguments).

Input: `=MIN(A2:A10, 0, C2:C4)`

Observed output: `0`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>MAX — value returned</summary>

Returns the largest numeric value from the supplied numbers or ranges, ignoring text, logical values, and blanks.

- number1 (number|range, required): First number, reference, or range that contains values to compare.
- number2 (number|range, optional): Optional additional number, reference, or range.
- rest (number|range, required): Optional extra numbers or ranges (up to 255 total arguments).

Input: `=MAX(A2:A10, 0, C2:C4)`

Observed output: `12`

Classification: value returned. Elapsed: 7 ms.

</details>

<details>
<summary>LINEST — Excel error</summary>

Returns statistics that describe a linear trend matching known data points, by fitting a straight line using the least squares method

- knownYs (value|range, required): Is the set of y-values you already know in the relationship y = mx + b
- knownXs (value|range, optional): Is an optional set of x-values that you may already know in the relationship y = mx + b
- constParam (value|range, optional): Is a logical value: the constant b is calculated normally if Const = TRUE or omitted; b is set equal to 0 if Const = FALSE
- stats (value|range, optional): Is a logical value: return additional regression statistics = TRUE; return m-coefficients and the constant b = FALSE or omitted

Input: `=LINEST(A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 11 ms.

</details>

<details>
<summary>TREND — Excel error</summary>

Returns numbers in a linear trend matching known data points, using the least squares method

- knownYs (value|range, required): Is a range or array of y-values you already know in the relationship y = mx + b
- knownXs (value|range, optional): Is an optional range or array of x-values that you know in the relationship y = mx + b, an array the same size as Known_y's
- newXs (value|range, optional): Is a range or array of new x-values for which you want TREND to return corresponding y-values
- constParam (value|range, optional): Is a logical value: the constant b is calculated normally if Const = TRUE or omitted; b is set equal to 0 if Const = FALSE

Input: `=TREND(A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 8 ms.

</details>

<details>
<summary>LOGEST — Excel error</summary>

Returns statistics that describe an exponential curve matching known data points

- knownYs (value|range, required): Is the set of y-values you already know in the relationship y = b*m^x
- knownXs (value|range, optional): Is an optional set of x-values that you may already know in the relationship y = b*m^x
- constParam (value|range, optional): Is a logical value: the constant b is calculated normally if Const = TRUE or omitted; b is set equal to 1 if Const = FALSE
- stats (value|range, optional): Is a logical value: return additional regression statistics = TRUE; return m-coefficients and the constant b = FALSE or omitted

Input: `=LOGEST(A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 12 ms.

</details>

<details>
<summary>GROWTH — Excel error</summary>

Returns numbers in an exponential growth trend matching known data points

- knownYs (value|range, required): Is the set of y-values you already know in the relationship y = b*m^x, an array or range of positive numbers
- knownXs (value|range, optional): Is an optional set of x-values that you may already know in the relationship y = b*m^x, an array or range the same size as Known_y's
- newXs (value|range, optional): Are new x-values for which you want GROWTH to return corresponding y-values
- constParam (value|range, optional): Is a logical value: the constant b is calculated normally if Const = TRUE; b is set equal to 1 if Const = FALSE or omitted

Input: `=GROWTH(A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 16 ms.

</details>

<details>
<summary>COUNTA — value returned</summary>

Counts the number of cells in a range that are not empty

- value1 (value|range, required): Are 1 to 255 arguments representing the values and cells you want to count. Values can be any type of information
- value2 (value|range, optional): Value argument for the calculation.
- rest (value|range, required): Value, reference, or range.

Input: `=COUNTA(A1:A5, C1:C5, B1:B5, C1:C5)`

Observed output: `20`

Classification: value returned. Elapsed: 18 ms.

</details>

<details>
<summary>MEDIAN — value returned</summary>

Returns the median, or the number in the middle of the set of given numbers

- number1 (number|range, required): Are 1 to 255 numbers or names, arrays, or references that contain numbers for which you want the median
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=MEDIAN(10, C1:C5, 20, C1:C5)`

Observed output: `9.5`

Classification: value returned. Elapsed: 16 ms.

</details>

<details>
<summary>FREQUENCY — value returned</summary>

Calculates how often values occur within a range of values and then returns a vertical array of numbers having one more element than Bins_array

- dataArray (array, required): Is an array of or reference to a set of values for which you want to count frequencies (blanks and text are ignored)
- binsArray (array, required): Is an array of or reference to intervals into which you want to group the values in data_array

Input: `=FREQUENCY({1,2;3,4}, {1,2;3,4})`

Observed output: `1`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>AVEDEV — value returned</summary>

Returns the average of the absolute deviations of data points from their mean. Arguments can be numbers or names, arrays, or references that contain numbers

- number1 (number|range, required): Are 1 to 255 arguments for which you want the average of the absolute deviations
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=AVEDEV(10, C1:C5, 20, C1:C5)`

Observed output: `4`

Classification: value returned. Elapsed: 15 ms.

</details>

<details>
<summary>GAMMALN — value returned</summary>

Returns the natural logarithm of the gamma function

- x (value|range, required): Is the value for which you want to calculate GAMMALN, a positive number

Input: `=GAMMALN(A1)`

Observed output: `0`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>FISHER — Excel error</summary>

Returns the Fisher transformation

- x (value|range, required): Is the value for which you want the transformation, a number between -1 and 1, excluding -1 and 1

Input: `=FISHER(A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 10 ms.

</details>

<details>
<summary>FISHERINV — value returned</summary>

Returns the inverse of the Fisher transformation: if y = FISHER(x), then FISHERINV(y) = x

- y (value|range, required): Is the value for which you want to perform the inverse of the transformation

Input: `=FISHERINV(A1)`

Observed output: `0.7615941559557649`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>STANDARDIZE — value returned</summary>

Returns a normalized value from a distribution characterized by a mean and standard deviation

- x (value|range, required): Is the value you want to normalize
- mean (value|range, required): Is the arithmetic mean of the distribution
- standardDev (value|range, required): Is the standard deviation of the distribution, a positive number

Input: `=STANDARDIZE(A1, A1, A1)`

Observed output: `0`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>PERMUT — value returned</summary>

Returns the number of permutations for a given number of objects that can be selected from the total objects

- numberParam (number|range, required): Is the total number of objects
- numberChosen (number|range, required): Is the number of objects in each permutation

Input: `=PERMUT(10, 10)`

Observed output: `3628800`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>CORREL — value returned</summary>

Returns the correlation coefficient between two data sets

- array1 (array, required): Is a cell range of values. The values should be numbers, names, arrays, or references that contain numbers
- array2 (array, required): Is a second cell range of values. The values should be numbers, names, arrays, or references that contain numbers

Input: `=CORREL({1,2;3,4}, {1,2;3,4})`

Observed output: `1.0000000000000002`

Classification: value returned. Elapsed: 22 ms.

</details>

<details>
<summary>INTERCEPT — Excel error</summary>

Calculates the point at which a line will intersect the y-axis by using a best-fit regression line plotted through the known x-values and y-values

- knownYs (value|range, required): Is the dependent set of observations or data and can be numbers or names, arrays, or references that contain numbers
- knownXs (value|range, required): Is the independent set of observations or data and can be numbers or names, arrays, or references that contain numbers

Input: `=INTERCEPT(A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 36 ms.

</details>

<details>
<summary>PEARSON — value returned</summary>

Returns the Pearson product moment correlation coefficient, r

- array1 (array, required): Is a set of independent values
- array2 (array, required): Is a set of dependent values

Input: `=PEARSON({1,2;3,4}, {1,2;3,4})`

Observed output: `1`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>RSQ — Excel error</summary>

Returns the square of the Pearson product moment correlation coefficient through the given data points

- knownYs (value|range, required): Is an array or range of data points and can be numbers or names, arrays, or references that contain numbers
- knownXs (value|range, required): Is an array or range of data points and can be numbers or names, arrays, or references that contain numbers

Input: `=RSQ(A1, A1)`

Observed output: `"#DIV/0!"`

Classification: Excel error. Elapsed: 14 ms.

</details>

<details>
<summary>STEYX — Excel error</summary>

Returns the standard error of the predicted y-value for each x in a regression

- knownYs (value|range, required): Is an array or range of dependent data points and can be numbers or names, arrays, or references that contain numbers
- knownXs (value|range, required): Is an array or range of independent data points and can be numbers or names, arrays, or references that contain numbers

Input: `=STEYX(A1, A1)`

Observed output: `"#DIV/0!"`

Classification: Excel error. Elapsed: 12 ms.

</details>

<details>
<summary>SLOPE — Excel error</summary>

Returns the slope of the linear regression line through the given data points

- knownYs (value|range, required): Is an array or cell range of numeric dependent data points and can be numbers or names, arrays, or references that contain numbers
- knownXs (value|range, required): Is the set of independent data points and can be numbers or names, arrays, or references that contain numbers

Input: `=SLOPE(A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 9 ms.

</details>

<details>
<summary>PROB — Excel error</summary>

Returns the probability that values in a range are between two limits or equal to a lower limit

- xRange (range, required): Is the range of numeric values of x with which there are associated probabilities
- probRange (range, required): Is the set of probabilities associated with values in X_range, values between 0 and 1 and excluding 0
- lowerLimit (value|range, required): Is the lower bound on the value for which you want a probability
- upperLimit (value|range, optional): Is the optional upper bound on the value. If omitted, PROB returns the probability that X_range values are equal to Lower_limit

Input: `=PROB(A1:A5, A1:A5, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 9 ms.

</details>

<details>
<summary>DEVSQ — value returned</summary>

Returns the sum of squares of deviations of data points from their sample mean

- number1 (number|range, required): Are 1 to 255 arguments, or an array or array reference, on which you want DEVSQ to calculate
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=DEVSQ(10, C1:C5, 20, C1:C5)`

Observed output: `290`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>GEOMEAN — value returned</summary>

Returns the geometric mean of an array or range of positive numeric data

- number1 (number|range, required): Are 1 to 255 numbers or names, arrays, or references that contain numbers for which you want the mean
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=GEOMEAN(10, C1:C5, 20, C1:C5)`

Observed output: `8.627475862383067`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>HARMEAN — value returned</summary>

Returns the harmonic mean of a data set of positive numbers: the reciprocal of the arithmetic mean of reciprocals

- number1 (number|range, required): Are 1 to 255 numbers or names, arrays, or references that contain numbers for which you want the harmonic mean
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=HARMEAN(10, C1:C5, 20, C1:C5)`

Observed output: `7.176079734219269`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>KURT — value returned</summary>

Returns the kurtosis of a data set

- number1 (number|range, required): Are 1 to 255 numbers or names, arrays, or references that contain numbers for which you want the kurtosis
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=KURT(10, C1:C5, 20, C1:C5)`

Observed output: `-0.27033055885850166`

Classification: value returned. Elapsed: 15 ms.

</details>

<details>
<summary>SKEW — value returned</summary>

Returns the skewness of a distribution: a characterization of the degree of asymmetry of a distribution around its mean

- number1 (number|range, required): Are 1 to 255 numbers or names, arrays, or references that contain numbers for which you want the skewness
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=SKEW(10, C1:C5, 20, C1:C5)`

Observed output: `0.36265452904062717`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>LARGE — value returned</summary>

Returns the k-th largest value in a data set. For example, the fifth largest number

- array (array, required): Is the array or range of data for which you want to determine the k-th largest value
- k (value|range, required): Is the position (from the largest) in the array or cell range of the value to return

Input: `=LARGE({1,2;3,4}, A1)`

Observed output: `4`

Classification: value returned. Elapsed: 17 ms.

</details>

<details>
<summary>SMALL — value returned</summary>

Returns the k-th smallest value in a data set. For example, the fifth smallest number

- array (array, required): Is an array or range of numerical data for which you want to determine the k-th smallest value
- k (value|range, required): Is the position (from the smallest) in the array or range of the value to return

Input: `=SMALL({1,2;3,4}, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>TRIMMEAN — Excel error</summary>

Returns the mean of the interior portion of a set of data values

- array (array, required): Is the range or array of values to trim and average
- percent (value|range, required): Is the fractional number of data points to exclude from the top and bottom of the data set

Input: `=TRIMMEAN({1,2;3,4}, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 11 ms.

</details>

<details>
<summary>COUNTIF — value returned</summary>

Counts the number of cells within a range that meet the given condition

- range (range, required): Is the range of cells from which you want to count nonblank cells
- criteria (criteria, required): Is the condition in the form of a number, expression, or text that defines which cells will be counted

Input: `=COUNTIF(A1:A5, ">0")`

Observed output: `5`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>COUNTBLANK — value returned</summary>

Counts the number of empty cells in a specified range of cells

- range (range, required): Is the range from which you want to count the empty cells

Input: `=COUNTBLANK(A1:A5)`

Observed output: `0`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>AVERAGEA — value returned</summary>

Returns the arithmetic mean while treating TRUE as 1, FALSE or text as 0, and numbers as-is.

- value1 (value|range, required): First value, reference, or range to evaluate (text and logical values are counted).
- value2 (value|range, optional): Additional value, reference, or range to include in the average.
- rest (value|range, required): Optional extra arguments (up to 255 total).

Input: `=AVERAGEA(A2:A5, TRUE, "5")`

Observed output: `3.3333333333333335`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>MAXA — value returned</summary>

Returns the largest value while treating TRUE as 1, FALSE or text as 0, and numbers as-is.

- value1 (value|range, required): First value, reference, or range to evaluate (text/logical values are counted).
- value2 (value|range, optional): Optional additional value, reference, or range.
- rest (value|range, required): Optional extra arguments (up to 255 total).

Input: `=MAXA(A2:A6, TRUE, "5")`

Observed output: `6`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>MINA — value returned</summary>

Returns the smallest value while treating TRUE as 1, FALSE or text as 0, and numbers as-is.

- value1 (value|range, required): First value, reference, or range to evaluate (text/logical values are counted).
- value2 (value|range, optional): Optional additional value, reference, or range.
- rest (value|range, required): Optional extra arguments (up to 255 total).

Input: `=MINA(A2:A6, FALSE, "5")`

Observed output: `0`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>STDEVPA — value returned</summary>

Calculates standard deviation based on an entire population, including logical values and text. Text and the logical value FALSE have the value 0; the logical value TRUE has the value 1

- value1 (value|range, required): Are 1 to 255 values corresponding to a population and can be values, names, arrays, or references that contain values
- value2 (value|range, optional): Value argument for the calculation.
- rest (value|range, required): Value, reference, or range.

Input: `=STDEVPA(A1:A5, C1:C5, B1:B5, C1:C5)`

Observed output: `4.205650960315181`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>VARPA — value returned</summary>

Calculates variance based on the entire population, including logical values and text. Text and the logical value FALSE have the value 0; the logical value TRUE has the value 1

- value1 (value|range, required): Are 1 to 255 value arguments corresponding to a population
- value2 (value|range, optional): Value argument for the calculation.
- rest (value|range, required): Value, reference, or range.

Input: `=VARPA(A1:A5, C1:C5, B1:B5, C1:C5)`

Observed output: `17.6875`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>STDEVA — value returned</summary>

Estimates standard deviation based on a sample, including logical values and text. Text and the logical value FALSE have the value 0; the logical value TRUE has the value 1

- value1 (value|range, required): Are 1 to 255 values corresponding to a sample of a population and can be values or names or references to values
- value2 (value|range, optional): Value argument for the calculation.
- rest (value|range, required): Value, reference, or range.

Input: `=STDEVA(A1:A5, C1:C5, B1:B5, C1:C5)`

Observed output: `4.314906841709515`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>VARA — value returned</summary>

Estimates variance based on a sample, including logical values and text. Text and the logical value FALSE have the value 0; the logical value TRUE has the value 1

- value1 (value|range, required): Are 1 to 255 value arguments corresponding to a sample of a population
- value2 (value|range, optional): Value argument for the calculation.
- rest (value|range, required): Value, reference, or range.

Input: `=VARA(A1:A5, C1:C5, B1:B5, C1:C5)`

Observed output: `18.61842105263158`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>COUNTIFS — value returned</summary>

Counts the number of cells specified by a given set of conditions or criteria

- criteriaRange (range, required): Is the range of cells you want evaluated for the particular condition
- criteria (criteria, required): Is the condition in the form of a number, expression, or text that defines which cells will be counted
- rest (value|range, required): Value, reference, or range.

Input: `=COUNTIFS(A1:A5, ">0", C1:C5, C1:C5)`

Observed output: `0`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>AVERAGEIF — Excel error</summary>

Calculates the mean of cells in a range that satisfy a single condition.

- range (range, required): Cells that are tested against the criteria.
- criteria (criteria, required): Condition such as ">0" or "apples" that determines which cells qualify.
- averageRange (range, optional): Optional cells to average when they differ from range; range is used when omitted.

Input: `=AVERAGEIF(A2:A10, ">100", B2:B10)`

Observed output: `"#DIV/0!"`

Classification: Excel error. Elapsed: 13 ms.

</details>

<details>
<summary>AVERAGEIFS — Excel error</summary>

Calculates the mean of cells that satisfy multiple range/criteria pairs.

- averageRange (range, required): Cells whose values are averaged when every criteria pair is satisfied.
- criteriaRange1 (range, required): Cells evaluated against the first criteria.
- criteria1 (criteria, required): First condition, such as "=North" or ">0".
- rest (range|criteria, required): Optional additional pairs supplied as criteria_range2, criteria2, criteria_range3, criteria3, and so on.

Input: `=AVERAGEIFS(B2:B10, A2:A10, "=West", C2:C10, ">0")`

Observed output: `"#DIV/0!"`

Classification: Excel error. Elapsed: 12 ms.

</details>

<details>
<summary>BINOM.DIST — value returned</summary>

Returns the individual term binomial distribution probability

- numberS (number|range, required): Is the number of successes in trials
- trials (value|range, required): Is the number of independent trials
- probabilityS (value|range, required): Is the probability of success on each trial
- cumulative (value|range, required): Is a logical value: for the cumulative distribution function, use TRUE; for the probability mass function, use FALSE

Input: `=BINOM.DIST(10, A1, A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>BINOM.INV — value returned</summary>

Returns the smallest value for which the cumulative binomial distribution is greater than or equal to a criterion value

- trials (value|range, required): Is the number of Bernoulli trials
- probabilityS (value|range, required): Is the probability of success on each trial, a number between 0 and 1 inclusive
- alpha (value|range, required): Is the criterion value, a number between 0 and 1 inclusive

Input: `=BINOM.INV(A1, A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>CONFIDENCE.NORM — value returned</summary>

Returns the confidence interval for a population mean, using a normal distribution

- alpha (value|range, required): Is the significance level used to compute the confidence level, a number greater than 0 and less than 1
- standardDev (value|range, required): Is the population standard deviation for the data range and is assumed to be known. Standard_dev must be greater than 0
- size (value|range, required): Is the sample size

Input: `=CONFIDENCE.NORM(A1, A1, A1)`

Observed output: `0`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>CONFIDENCE.T — value returned</summary>

Returns the confidence interval for a population mean, using a Student's T distribution

- alpha (value|range, required): Is the significance level used to compute the confidence level, a number greater than 0 and less than 1
- standardDev (value|range, required): Is the population standard deviation for the data range and is assumed to be known. Standard_dev must be greater than 0
- size (value|range, required): Is the sample size

Input: `=CONFIDENCE.T(A1, A1, A1)`

Observed output: `0`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>CHISQ.TEST — value returned</summary>

Returns the test for independence: the value from the chi-squared distribution for the statistic and the appropriate degrees of freedom

- actualRange (range, required): Is the range of data that contains observations to test against expected values
- expectedRange (range, required): Is the range of data that contains the ratio of the product of row totals and column totals to the grand total

Input: `=CHISQ.TEST(A1:A5, A1:A5)`

Observed output: `1`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>F.TEST — Excel error</summary>

Returns the result of an F-test, the two-tailed probability that the variances in Array1 and Array2 are not significantly different

- array1 (array, required): Is the first array or range of data and can be numbers or names, arrays, or references that contain numbers (blanks are ignored)
- array2 (array, required): Is the second array or range of data and can be numbers or names, arrays, or references that contain numbers (blanks are ignored)

Input: `=F.TEST({1,2;3,4}, {1,2;3,4})`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 9 ms.

</details>

<details>
<summary>COVARIANCE.P — value returned</summary>

Returns population covariance, the average of the products of deviations for each data point pair in two data sets

- array1 (array, required): Is the first cell range of integers and must be numbers, arrays, or references that contain numbers
- array2 (array, required): Is the second cell range of integers and must be numbers, arrays, or references that contain numbers

Input: `=COVARIANCE.P({1,2;3,4}, {1,2;3,4})`

Observed output: `1.25`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>COVARIANCE.S — value returned</summary>

Returns sample covariance, the average of the products of deviations for each data point pair in two data sets

- array1 (array, required): Is the first cell range of integers and must be numbers, arrays, or references that contain numbers
- array2 (array, required): Is the second cell range of integers and must be numbers, arrays, or references that contain numbers

Input: `=COVARIANCE.S({1,2;3,4}, {1,2;3,4})`

Observed output: `1.6666666666666667`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>EXPON.DIST — value returned</summary>

Returns the exponential distribution

- x (value|range, required): Is the value of the function, a nonnegative number
- lambda (value|range, required): Is the parameter value, a positive number
- cumulative (value|range, required): Is a logical value for the function to return: the cumulative distribution function = TRUE; the probability density function = FALSE

Input: `=EXPON.DIST(A1, A1, A1)`

Observed output: `0.6321205588285577`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>GAMMA.DIST — value returned</summary>

Returns the gamma distribution

- x (value|range, required): Is the value at which you want to evaluate the distribution, a nonnegative number
- alpha (value|range, required): Is a parameter to the distribution, a positive number
- beta (value|range, required): Is a parameter to the distribution, a positive number. If beta = 1, GAMMA.DIST returns the standard gamma distribution
- cumulative (value|range, required): Is a logical value: return the cumulative distribution function = TRUE; return the probability mass function = FALSE or omitted

Input: `=GAMMA.DIST(A1, A1, A1, A1)`

Observed output: `0.6321205588285578`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>GAMMA.INV — Excel error</summary>

Returns the inverse of the gamma cumulative distribution: if p = GAMMA.DIST(x,...), then GAMMA.INV(p,...) = x

- probability (value|range, required): Is the probability associated with the gamma distribution, a number between 0 and 1, inclusive
- alpha (value|range, required): Is a parameter to the distribution, a positive number
- beta (value|range, required): Is a parameter to the distribution, a positive number. If beta = 1, GAMMA.INV returns the inverse of the standard gamma distribution

Input: `=GAMMA.INV(A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 9 ms.

</details>

<details>
<summary>MODE.MULT — value returned</summary>

Returns a vertical array of the most frequently occurring, or repetitive, values in an array or range of data. For a horizontal array, use =TRANSPOSE(MODE.MULT(number1,number2,...))

- number1 (number|range, required): Are 1 to 255 numbers, or names, arrays, or references that contain numbers for which you want the mode
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=MODE.MULT(10, C1:C5, 20, C1:C5)`

Observed output: `3`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>MODE.SNGL — value returned</summary>

Returns the most frequently occurring, or repetitive, value in an array or range of data

- number1 (number|range, required): Are 1 to 255 numbers, or names, arrays, or references that contain numbers for which you want the mode
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=MODE.SNGL(10, C1:C5, 20, C1:C5)`

Observed output: `3`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>NORM.DIST — value returned</summary>

Returns the normal distribution for the specified mean and standard deviation

- x (value|range, required): Is the value for which you want the distribution
- mean (value|range, required): Is the arithmetic mean of the distribution
- standardDev (value|range, required): Is the standard deviation of the distribution, a positive number
- cumulative (value|range, required): Is a logical value: for the cumulative distribution function, use TRUE; for the probability density function, use FALSE

Input: `=NORM.DIST(A1, A1, A1, A1)`

Observed output: `0.5`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>NORM.INV — value returned</summary>

Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation

- probability (value|range, required): Is a probability corresponding to the normal distribution, a number between 0 and 1 inclusive
- mean (value|range, required): Is the arithmetic mean of the distribution
- standardDev (value|range, required): Is the standard deviation of the distribution, a positive number

Input: `=NORM.INV(A1, A1, A1)`

Observed output: `142.4213562373095`

Classification: value returned. Elapsed: 26 ms.

</details>

<details>
<summary>PERCENTILE.EXC — Excel error</summary>

Returns the k-th percentile of values in a range, where k is in the range 0..1, exclusive

- array (array, required): Is the array or range of data that defines relative standing
- k (value|range, required): Is the percentile value that is between 0 through 1, inclusive

Input: `=PERCENTILE.EXC({1,2;3,4}, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 15 ms.

</details>

<details>
<summary>PERCENTILE.INC — value returned</summary>

Returns the k-th percentile of values in a range, where k is in the range 0..1, inclusive

- array (array, required): Is the array or range of data that defines relative standing
- k (value|range, required): Is the percentile value that is between 0 through 1, inclusive

Input: `=PERCENTILE.INC({1,2;3,4}, A1)`

Observed output: `4`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>PERCENTRANK.EXC — Excel error</summary>

Returns the rank of a value in a data set as a percentage (0..1, exclusive) of the data set

- array (array, required): Is the array or range of data with numeric values that defines relative standing
- x (value|range, required): Is the value for which you want to know the rank
- significance (value|range, optional): Is an optional value that identifies the number of significant digits for the returned percentage, three digits if omitted (0.xxx%)

Input: `=PERCENTRANK.EXC({1,2;3,4}, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 16 ms.

</details>

<details>
<summary>PERCENTRANK.INC — value returned</summary>

Returns the rank of a value in a data set as a percentage (0..1, inclusive) of the data set

- array (array, required): Is the array or range of data with numeric values that defines relative standing
- x (value|range, required): Is the value for which you want to know the rank
- significance (value|range, optional): Is an optional value that identifies the number of significant digits for the returned percentage, three digits if omitted (0.xxx%)

Input: `=PERCENTRANK.INC({1,2;3,4}, A1, A1)`

Observed output: `0`

Classification: value returned. Elapsed: 21 ms.

</details>

<details>
<summary>POISSON.DIST — value returned</summary>

Returns the Poisson distribution

- x (value|range, required): Is the number of events
- mean (value|range, required): Is the expected numeric value, a positive number
- cumulative (value|range, required): Is a logical value: for the cumulative Poisson probability, use TRUE; for the Poisson probability mass function, use FALSE

Input: `=POISSON.DIST(A1, A1, A1)`

Observed output: `0.7357588823428847`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>QUARTILE.EXC — value returned</summary>

Returns the quartile of a data set, based on percentile values from 0..1, exclusive

- array (array, required): Is the array or cell range of numeric values for which you want the quartile value
- quart (value|range, required): Is a number: minimum value = 0; 1st quartile = 1; median value = 2; 3rd quartile = 3; maximum value = 4

Input: `=QUARTILE.EXC({1,2;3,4}, A1)`

Observed output: `1.25`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>QUARTILE.INC — value returned</summary>

Returns the quartile of a data set, based on percentile values from 0..1, inclusive

- array (array, required): Is the array or cell range of numeric values for which you want the quartile value
- quart (value|range, required): Is a number: minimum value = 0; 1st quartile = 1; median value = 2; 3rd quartile = 3; maximum value = 4

Input: `=QUARTILE.INC({1,2;3,4}, A1)`

Observed output: `1.75`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>RANK.AVG — value returned</summary>

Returns the rank of a number in a list of numbers: its size relative to other values in the list; if more than one value has the same rank, the average rank is returned

- numberParam (number|range, required): Is the number for which you want to find the rank
- ref (value|range, required): Is an array of, or a reference to, a list of numbers. Nonnumeric values are ignored
- order (value|range, optional): Is a number: rank in the list sorted descending = 0 or omitted; rank in the list sorted ascending = any nonzero value

Input: `=RANK.AVG(10, A1, A1)`

Observed output: `0`

Classification: value returned. Elapsed: 18 ms.

</details>

<details>
<summary>RANK.EQ — value returned</summary>

Returns the rank of a number in a list of numbers: its size relative to other values in the list; if more than one value has the same rank, the top rank of that set of values is returned

- numberParam (number|range, required): Is the number for which you want to find the rank
- ref (value|range, required): Is an array of, or a reference to, a list of numbers. Nonnumeric values are ignored
- order (value|range, optional): Is a number: rank in the list sorted descending = 0 or omitted; rank in the list sorted ascending = any nonzero value

Input: `=RANK.EQ(10, A1, A1)`

Observed output: `0`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>STDEV.S — value returned</summary>

Estimates standard deviation based on a sample (ignores logical values and text in the sample)

- number1 (number|range, required): Are 1 to 255 numbers corresponding to a sample of a population and can be numbers or references that contain numbers
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=STDEV.S(10, C1:C5, 20, C1:C5)`

Observed output: `5.134553180524705`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>STDEV.P — value returned</summary>

Calculates standard deviation based on the entire population given as arguments (ignores logical values and text)

- number1 (number|range, required): Are 1 to 255 numbers corresponding to a population and can be numbers or references that contain numbers
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=STDEV.P(10, C1:C5, 20, C1:C5)`

Observed output: `4.915960401250875`

Classification: value returned. Elapsed: 15 ms.

</details>

<details>
<summary>T.DIST — value returned</summary>

Returns the left-tailed Student's t-distribution

- x (value|range, required): Is the numeric value at which to evaluate the distribution
- degFreedom (value|range, required): Is an integer indicating the number of degrees of freedom that characterize the distribution
- cumulative (value|range, required): Is a logical value: for the cumulative distribution function, use TRUE; for the probability density function, use FALSE

Input: `=T.DIST(A1, A1, A1)`

Observed output: `0.24999999852683763`

Classification: value returned. Elapsed: 17 ms.

</details>

<details>
<summary>T.DIST.2T — exception-like text</summary>

Returns the two-tailed Student's t-distribution

- x (value|range, required): Is the numeric value at which to evaluate the distribution
- degFreedom (value|range, required): Is an integer indicating the number of degrees of freedom that characterize the distribution

Input: `=T.DIST.2T(A1, A1)`

Observed output: `"XDe._2T is not a function"`

Classification: exception-like text. Elapsed: 14 ms.

</details>

<details>
<summary>T.DIST.RT — value returned</summary>

Returns the right-tailed Student's t-distribution

- x (value|range, required): Is the numeric value at which to evaluate the distribution
- degFreedom (value|range, required): Is an integer indicating the number of degrees of freedom that characterize the distribution

Input: `=T.DIST.RT(A1, A1)`

Observed output: `0.24999999852683763`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>T.INV — Excel error</summary>

Returns the left-tailed inverse of the Student's t-distribution

- probability (value|range, required): Is the probability associated with the two-tailed Student's t-distribution, a number between 0 and 1 inclusive
- degFreedom (value|range, required): Is a positive integer indicating the number of degrees of freedom to characterize the distribution

Input: `=T.INV(A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 15 ms.

</details>

<details>
<summary>T.INV.2T — Excel error</summary>

Returns the two-tailed inverse of the Student's t-distribution

- probability (value|range, required): Is the probability associated with the two-tailed Student's t-distribution, a number between 0 and 1 inclusive
- degFreedom (value|range, required): Is a positive integer indicating the number of degrees of freedom to characterize the distribution

Input: `=T.INV.2T(A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 10 ms.

</details>

<details>
<summary>VAR.S — value returned</summary>

Estimates variance based on a sample (ignores logical values and text in the sample)

- number1 (number|range, required): Are 1 to 255 numeric arguments corresponding to a sample of a population
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=VAR.S(10, C1:C5, 20, C1:C5)`

Observed output: `26.363636363636363`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>VAR.P — value returned</summary>

Calculates variance based on the entire population (ignores logical values and text in the population)

- number1 (number|range, required): Are 1 to 255 numeric arguments corresponding to a population
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=VAR.P(10, C1:C5, 20, C1:C5)`

Observed output: `24.166666666666668`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>WEIBULL.DIST — value returned</summary>

Returns the Weibull distribution

- x (value|range, required): Is the value at which to evaluate the function, a nonnegative number
- alpha (value|range, required): Is a parameter to the distribution, a positive number
- beta (value|range, required): Is a parameter to the distribution, a positive number
- cumulative (value|range, required): Is a logical value: for the cumulative distribution function, use TRUE; for the probability mass function, use FALSE

Input: `=WEIBULL.DIST(A1, A1, A1, A1)`

Observed output: `0.6321205588285577`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>BETA.DIST — value returned</summary>

Returns the beta probability distribution function

- x (value|range, required): Is the value between A and B at which to evaluate the function
- alpha (value|range, required): Is a parameter to the distribution and must be greater than 0
- beta (value|range, required): Is a parameter to the distribution and must be greater than 0
- cumulative (value|range, required): Is a logical value: for the cumulative distribution function, use TRUE; for the probability density function, use FALSE
- a (value|range, optional): Is an optional lower bound to the interval of x. If omitted, A = 0
- b (value|range, optional): Is an optional upper bound to the interval of x. If omitted, B = 1

Input: `=BETA.DIST(A1, A1, A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>BETA.INV — value returned</summary>

Returns the inverse of the cumulative beta probability density function (BETA.DIST)

- probability (value|range, required): Is a probability associated with the beta distribution
- alpha (value|range, required): Is a parameter to the distribution and must be greater than 0
- beta (value|range, required): Is a parameter to the distribution and must be greater than 0
- a (value|range, optional): Is an optional lower bound to the interval of x. If omitted, A = 0
- b (value|range, optional): Is an optional upper bound to the interval of x. If omitted, B = 1

Input: `=BETA.INV(A1, A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>CHISQ.DIST — value returned</summary>

Returns the left-tailed probability of the chi-squared distribution

- x (value|range, required): Is the value at which you want to evaluate the distribution, a nonnegative number
- degFreedom (value|range, required): Is the number of degrees of freedom, a number between 1 and 10^10, excluding 10^10
- cumulative (value|range, required): Is a logical value for the function to return: the cumulative distribution function = TRUE; the probability density function = FALSE

Input: `=CHISQ.DIST(A1, A1, A1)`

Observed output: `0.6826894921370564`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>CHISQ.DIST.RT — value returned</summary>

Returns the right-tailed probability of the chi-squared distribution

- x (value|range, required): Is the value at which you want to evaluate the distribution, a nonnegative number
- degFreedom (value|range, required): Is the number of degrees of freedom, a number between 1 and 10^10, excluding 10^10

Input: `=CHISQ.DIST.RT(A1, A1)`

Observed output: `0.31731050786294357`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>CHISQ.INV — Excel error</summary>

Returns the inverse of the left-tailed probability of the chi-squared distribution

- probability (value|range, required): Is a probability associated with the chi-squared distribution, a value between 0 and 1 inclusive
- degFreedom (value|range, required): Is the number of degrees of freedom, a number between 1 and 10^10, excluding 10^10

Input: `=CHISQ.INV(A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 9 ms.

</details>

<details>
<summary>CHISQ.INV.RT — Excel error</summary>

Returns the inverse of the right-tailed probability of the chi-squared distribution

- probability (value|range, required): Is a probability associated with the chi-squared distribution, a value between 0 and 1 inclusive
- degFreedom (value|range, required): Is the number of degrees of freedom, a number between 1 and 10^10, excluding 10^10

Input: `=CHISQ.INV.RT(A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 8 ms.

</details>

<details>
<summary>F.DIST — value returned</summary>

Returns the (left-tailed) F probability distribution (degree of diversity) for two data sets

- x (value|range, required): Is the value at which to evaluate the function, a nonnegative number
- degFreedom1 (value|range, required): Is the numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10
- degFreedom2 (value|range, required): Is the denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10
- cumulative (value|range, required): Is a logical value for the function to return: the cumulative distribution function = TRUE; the probability density function = FALSE

Input: `=F.DIST(A1, A1, A1, A1)`

Observed output: `0.5000000180073508`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>F.DIST.RT — value returned</summary>

Returns the (right-tailed) F probability distribution (degree of diversity) for two data sets

- x (value|range, required): Is the value at which to evaluate the function, a nonnegative number
- degFreedom1 (value|range, required): Is the numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10
- degFreedom2 (value|range, required): Is the denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10

Input: `=F.DIST.RT(A1, A1, A1)`

Observed output: `0.49999998199264917`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>F.INV — Excel error</summary>

Returns the inverse of the (left-tailed) F probability distribution: if p = F.DIST(x,...), then F.INV(p,...) = x

- probability (value|range, required): Is a probability associated with the F cumulative distribution, a number between 0 and 1 inclusive
- degFreedom1 (value|range, required): Is the numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10
- degFreedom2 (value|range, required): Is the denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10

Input: `=F.INV(A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 7 ms.

</details>

<details>
<summary>F.INV.RT — Excel error</summary>

Returns the inverse of the (right-tailed) F probability distribution: if p = F.DIST.RT(x,...), then F.INV.RT(p,...) = x

- probability (value|range, required): Is a probability associated with the F cumulative distribution, a number between 0 and 1 inclusive
- degFreedom1 (value|range, required): Is the numerator degrees of freedom, a number between 1 and 10^10, excluding 10^10
- degFreedom2 (value|range, required): Is the denominator degrees of freedom, a number between 1 and 10^10, excluding 10^10

Input: `=F.INV.RT(A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 11 ms.

</details>

<details>
<summary>HYPGEOM.DIST — value returned</summary>

Returns the hypergeometric distribution

- sampleS (value|range, required): Is the number of successes in the sample
- numberSample (number|range, required): Is the size of the sample
- populationS (value|range, required): Is the number of successes in the population
- numberPop (number|range, required): Is the population size
- cumulative (value|range, required): Is a logical value: for the cumulative distribution function, use TRUE; for the probability density function, use FALSE

Input: `=HYPGEOM.DIST(A1, 10, A1, 10)`

Observed output: `1`

Classification: value returned. Elapsed: 7 ms.

</details>

<details>
<summary>LOGNORM.DIST — value returned</summary>

Returns the lognormal distribution of x, where ln(x) is normally distributed with parameters Mean and Standard_dev

- x (value|range, required): Is the value at which to evaluate the function, a positive number
- mean (value|range, required): Is the mean of ln(x)
- standardDev (value|range, required): Is the standard deviation of ln(x), a positive number
- cumulative (value|range, required): Is a logical value: for the cumulative distribution function, use TRUE; for the probability density function, use FALSE

Input: `=LOGNORM.DIST(A1, A1, A1, A1)`

Observed output: `0.15865525393145707`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>LOGNORM.INV — value returned</summary>

Returns the inverse of the lognormal cumulative distribution function of x, where ln(x) is normally distributed with parameters Mean and Standard_dev

- probability (value|range, required): Is a probability associated with the lognormal distribution, a number between 0 and 1, inclusive
- mean (value|range, required): Is the mean of ln(x)
- standardDev (value|range, required): Is the standard deviation of ln(x), a positive number

Input: `=LOGNORM.INV(A1, A1, A1)`

Observed output: `7.125397860548684e+61`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>NEGBINOM.DIST — value returned</summary>

Returns the negative binomial distribution, the probability that there will be Number_f failures before the Number_s-th success, with Probability_s probability of a success

- numberF (number|range, required): Is the number of failures
- numberS (number|range, required): Is the threshold number of successes
- probabilityS (value|range, required): Is the probability of a success; a number between 0 and 1
- cumulative (value|range, required): Is a logical value: for the cumulative distribution function, use TRUE; for the probability mass function, use FALSE

Input: `=NEGBINOM.DIST(10, 10, A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>NORM.S.DIST — value returned</summary>

Returns the standard normal distribution (has a mean of zero and a standard deviation of one)

- z (value|range, required): Is the value for which you want the distribution
- cumulative (value|range, required): Is a logical value for the function to return: the cumulative distribution function = TRUE; the probability density function = FALSE

Input: `=NORM.S.DIST(A1, A1)`

Observed output: `0.8413447460685429`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>NORM.S.INV — value returned</summary>

Returns the inverse of the standard normal cumulative distribution (has a mean of zero and a standard deviation of one)

- probability (value|range, required): Is a probability corresponding to the normal distribution, a number between 0 and 1 inclusive

Input: `=NORM.S.INV(A1)`

Observed output: `141.4213562373095`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>T.TEST — value returned</summary>

Returns the probability associated with a Student's t-Test

- array1 (array, required): Is the first data set
- array2 (array, required): Is the second data set
- tails (value|range, required): Specifies the number of distribution tails to return: one-tailed distribution = 1; two-tailed distribution = 2
- typeParam (value|range, required): Is the kind of t-test: paired = 1, two-sample equal variance (homoscedastic) = 2, two-sample unequal variance = 3

Input: `=T.TEST({1,2;3,4}, {1,2;3,4}, A1, A1)`

Observed output: `1`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>Z.TEST — value returned</summary>

Returns the one-tailed P-value of a z-test

- array (array, required): Is the array or range of data against which to test X
- x (value|range, required): Is the value to test
- sigma (value|range, optional): Is the population (known) standard deviation. If omitted, the sample standard deviation is used

Input: `=Z.TEST({1,2;3,4}, A1, A1)`

Observed output: `0.0013498980316301035`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>GAMMALN.PRECISE — value returned</summary>

Returns the natural logarithm of the gamma function

- x (value|range, required): Is the value for which you want to calculate GAMMALN.PRECISE, a positive number

Input: `=GAMMALN.PRECISE(A1)`

Observed output: `0`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>PERMUTATIONA — value returned</summary>

Returns the number of permutations for a given number of objects (with repetitions) that can be selected from the total objects

- numberParam (number|range, required): Is the total number of objects
- numberChosen (number|range, required): Is the number of objects in each permutation

Input: `=PERMUTATIONA(10, 10)`

Observed output: `10000000000`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>BINOM.DIST.RANGE — Excel error</summary>

Returns the probability of a trial result using a binomial distribution

- trials (value|range, required): Is the number of independent trials
- probabilityS (value|range, required): Is the probability of success on each trial
- numberS (number|range, required): Is the number of successes in trials
- numberS2 (number|range, optional): If provided this function returns the probability that the number of successful trials shall lie between number_s and number_s2

Input: `=BINOM.DIST.RANGE(A1, A1, 10)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 13 ms.

</details>

<details>
<summary>GAMMA — value returned</summary>

Returns the Gamma function value

- x (value|range, required): Is the value for which you want to calculate Gamma

Input: `=GAMMA(A1)`

Observed output: `1`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>SKEW.P — value returned</summary>

Returns the skewness of a distribution based on a population: a characterization of the degree of asymmetry of a distribution around its mean

- number1 (number|range, required): Are 1 to 254 numbers or names, arrays, or references that contain numbers for which you want the population skewness
- number2 (number|range, optional): Additional number, range, or reference to include.
- rest (value|range, required): Value, reference, or range.

Input: `=SKEW.P(10, C1:C5, 20, C1:C5)`

Observed output: `0.31565025168554967`

Classification: value returned. Elapsed: 27 ms.

</details>

<details>
<summary>GAUSS — value returned</summary>

Returns 0.5 less than the standard normal cumulative distribution

- x (value|range, required): Is the value for which you want the distribution

Input: `=GAUSS(A1)`

Observed output: `0.3413447460685429`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>PHI — value returned</summary>

Returns the value of the density function for a standard normal distribution

- x (value|range, required): Is the number for which you want the density of the standard normal distribution

Input: `=PHI(A1)`

Observed output: `0.24197072451914337`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>FORECAST.LINEAR — Excel error</summary>

Calculates, or predicts, a future value along a linear trend by using existing values

- x (value|range, required): Is the data point for which you want to predict a value and must be a numeric value
- knownYs (value|range, required): Is the dependent array or range of numeric data
- knownXs (value|range, required): Is the independent array or range of numeric data. The variance of Known_x's must not be zero

Input: `=FORECAST.LINEAR(A1, A1, A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 8 ms.

</details>

<details>
<summary>MAXIFS — value returned</summary>

Returns the maximum value from maxRange after applying one or more criteria ranges.

- maxRange (range, required): Cells that contain the values to compare.
- criteriaRange1 (range, required): First range to test against the paired criteria.
- criteria1 (criteria, required): First condition, such as "=East" or ">0".
- rest (range|criteria, required): Optional additional pairs supplied as criteria_range2, criteria2, criteria_range3, criteria3, and so on.

Input: `=MAXIFS(B2:B20, A2:A20, "North", C2:C20, ">1000")`

Observed output: `0`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>MINIFS — value returned</summary>

Returns the minimum value from minRange after applying one or more criteria ranges.

- minRange (range, required): Cells that contain the values to compare.
- criteriaRange1 (range, required): First range to test against the paired criteria.
- criteria1 (criteria, required): First condition, such as "=East" or ">0".
- rest (range|criteria, required): Optional additional pairs supplied as criteria_range2, criteria2, criteria_range3, criteria3, and so on.

Input: `=MINIFS(B2:B20, A2:A20, "North", C2:C20, ">0")`

Observed output: `0`

Classification: value returned. Elapsed: 14 ms.

</details>


### Formula category: text

<details>
<summary>DOLLAR — value returned</summary>

Converts a number to text, using currency format

- numberParam (number|range, required): Is a number, a reference to a cell containing a number, or a formula that evaluates to a number
- decimals (value|range, optional): Is the number of digits to the right of the decimal point. The number is rounded as necessary; if omitted, Decimals = 2

Input: `=DOLLAR(10, A1)`

Observed output: `"$10.0"`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>FIXED — value returned</summary>

Rounds a number to the specified number of decimals and returns the result as text with or without commas

- numberParam (number|range, required): Is the number you want to round and convert to text
- decimals (value|range, optional): Is the number of digits to the right of the decimal point. If omitted, Decimals = 2
- noCommas (value|range, optional): Is a logical value: do not display commas in the returned text = TRUE; do display commas in the returned text = FALSE or omitted

Input: `=FIXED(10, A1)`

Observed output: `"10.0"`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>REPT — exception-like text</summary>

Repeats text a given number of times. Use REPT to fill a cell with a number of instances of a text string

- text (string, required): Is the text you want to repeat
- numberTimes (time, required): Is a positive number specifying the number of times to repeat text

Input: `=REPT("text", TIME(9, 0, 0))`

Observed output: `"Invalid array length"`

Classification: exception-like text. Elapsed: 8 ms.

</details>

<details>
<summary>MID — value returned</summary>

Returns the characters from the middle of a text string, given a starting position and length

- text (string, required): Is the text string from which you want to extract the characters
- startNum (value|range, required): Is the position of the first character you want to extract. The first character in Text is 1
- numChars (value|range, required): Specifies how many characters to return from Text

Input: `=MID("text", A1, A1)`

Observed output: `"t"`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>LEN — value returned</summary>

Returns the number of characters in a text string

- text (string, required): Is the text whose length you want to find. Spaces count as characters

Input: `=LEN("text")`

Observed output: `4`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>VALUE — value returned</summary>

Converts a text string that represents a number to a number

- text (string, required): Is the text enclosed in quotation marks or a reference to a cell containing the text you want to convert

Input: `=VALUE("text")`

Observed output: `0`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>TEXT — value returned</summary>

Converts a value to text in a specific number format

- value (value|range, required): Is a number, a formula that evaluates to a numeric value, or a reference to a cell containing a numeric value
- formatText (string, required): Is a number format in text form from the Category box on the Number tab in the Format Cells dialog box

Input: `=TEXT(A1:A5, "#,##0.00")`

Observed output: `"1.00"`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>SEARCH — value returned</summary>

Returns the number of the character at which a specific character or text string is first found, reading left to right (not case-sensitive)

- findText (string, required): Is the text you want to find. You can use the ? and * wildcard characters; use ~? and ~* to find the ? and * characters
- withinText (string, required): Is the text in which you want to search for Find_text
- startNum (value|range, optional): Is the character number in Within_text, counting from the left, at which you want to start searching. If omitted, 1 is used

Input: `=SEARCH("text", "text", A1)`

Observed output: `1`

Classification: value returned. Elapsed: 14 ms.

</details>

<details>
<summary>CHAR — value returned</summary>

Returns the character specified by the code number from the character set for your computer

- numberParam (number|range, required): Is a number between 1 and 255 specifying which character you want

Input: `=CHAR(10)`

Observed output: `"\n"`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>LOWER — value returned</summary>

Converts all letters in a text string to lowercase

- text (string, required): Is the text you want to convert to lowercase. Characters in Text that are not letters are not changed

Input: `=LOWER("text")`

Observed output: `"text"`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>UPPER — value returned</summary>

Converts a text string to all uppercase letters

- text (string, required): Is the text you want converted to uppercase, a reference or a text string

Input: `=UPPER("text")`

Observed output: `"TEXT"`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>PROPER — value returned</summary>

Converts a text string to proper case; the first letter in each word to uppercase, and all other letters to lowercase

- text (string, required): Is text enclosed in quotation marks, a formula that returns text, or a reference to a cell containing text to partially capitalize

Input: `=PROPER("text")`

Observed output: `"Text"`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>LEFT — value returned</summary>

Returns the specified number of characters from the start of a text string

- text (string, required): Is the text string containing the characters you want to extract
- numChars (value|range, optional): Specifies how many characters you want LEFT to extract; 1 if omitted

Input: `=LEFT("text", A1)`

Observed output: `"t"`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>RIGHT — value returned</summary>

Returns the specified number of characters from the end of a text string

- text (string, required): Is the text string that contains the characters you want to extract
- numChars (value|range, optional): Specifies how many characters you want to extract, 1 if omitted

Input: `=RIGHT("text", A1)`

Observed output: `"t"`

Classification: value returned. Elapsed: 7 ms.

</details>

<details>
<summary>EXACT — value returned</summary>

Checks whether two text strings are exactly the same, and returns TRUE or FALSE. EXACT is case-sensitive

- text1 (string, required): Is the first text string
- text2 (string, required): Is the second text string

Input: `=EXACT("text1", "text2")`

Observed output: `false`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>TRIM — value returned</summary>

Removes all spaces from a text string except for single spaces between words

- text (string, required): Is the text from which you want spaces removed

Input: `=TRIM("text")`

Observed output: `"text"`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>REPLACE — value returned</summary>

Replaces part of a text string with a different text string

- oldText (string, required): Is text in which you want to replace some characters
- startNum (value|range, required): Is the position of the character in Old_text that you want to replace with New_text
- numChars (value|range, required): Is the number of characters in Old_text that you want to replace
- newText (string, required): Is the text that will replace characters in Old_text

Input: `=REPLACE("text", A1, A1, "text")`

Observed output: `"textext"`

Classification: value returned. Elapsed: 13 ms.

</details>

<details>
<summary>SUBSTITUTE — value returned</summary>

Replaces existing text with new text in a text string

- text (string, required): Is the text or the reference to a cell containing text in which you want to substitute characters
- oldText (string, required): Is the existing text you want to replace. If the case of Old_text does not match the case of text, SUBSTITUTE will not replace the text
- newText (string, required): Is the text you want to replace Old_text with
- instanceNum (value|range, optional): Specifies which occurrence of Old_text you want to replace. If omitted, every instance of Old_text is replaced

Input: `=SUBSTITUTE("text", "text", "text")`

Observed output: `"text"`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>CODE — value returned</summary>

Returns a numeric code for the first character in a text string, in the character set used by your computer

- text (string, required): Is the text for which you want the code of the first character

Input: `=CODE("text")`

Observed output: `116`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>FIND — value returned</summary>

Returns the starting position of one text string within another text string. FIND is case-sensitive

- findText (string, required): Is the text you want to find. Use double quotes (empty text) to match the first character in Within_text; wildcard characters not allowed
- withinText (string, required): Is the text containing the text you want to find
- startNum (value|range, optional): Specifies the character at which to start the search. The first character in Within_text is character number 1. If omitted, Start_num = 1

Input: `=FIND("text", "text", A1)`

Observed output: `1`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>T — value returned</summary>

Checks whether a value is text, and returns the text if it is, or returns double quotes (empty text) if it is not

- value (value|range, required): Is the value to test

Input: `=T(A1:A5)`

Observed output: `""`

Classification: value returned. Elapsed: 8 ms.

</details>

<details>
<summary>CLEAN — value returned</summary>

Removes all nonprintable characters from text

- text (string, required): Is any worksheet information from which you want to remove nonprintable characters

Input: `=CLEAN("text")`

Observed output: `"text"`

Classification: value returned. Elapsed: 11 ms.

</details>

<details>
<summary>BAHTTEXT — unimplemented text</summary>

Converts a number to text (baht)

- numberParam (number|range, required): Is a number that you want to convert

Input: `=BAHTTEXT(10)`

Observed output: `"BAHTTEXT is not implemented. numberParam=10"`

Classification: unimplemented text. Elapsed: 12 ms.

</details>

<details>
<summary>UNICHAR — value returned</summary>

Returns the Unicode character referenced by the given numeric value

- numberParam (number|range, required): Is the Unicode number representing a character

Input: `=UNICHAR(10)`

Observed output: `"\n"`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>UNICODE — value returned</summary>

Returns the number (code point) corresponding to the first character of the text

- text (string, required): Is the character that you want the Unicode value of

Input: `=UNICODE("text")`

Observed output: `116`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>NUMBERVALUE — Excel error</summary>

Converts text to number in a locale-independent manner

- text (string, required): Is the string representing the number you want to convert
- decimalSeparator (value|range, optional): Is the character used as the decimal separator in the string
- groupSeparator (value|range, optional): Is the character used as the group separator in the string

Input: `=NUMBERVALUE("text", A1)`

Observed output: `"#NUM!"`

Classification: Excel error. Elapsed: 15 ms.

</details>

<details>
<summary>TEXTJOIN — value returned</summary>

Concatenates a list or range of text strings using a delimiter

- delimiter (string, required): Character or string to insert between each text item
- ignoreEmpty (value|range, required): If TRUE(default), ignores empty cells
- text1 (string, required): Are 1 to 252 text strings or ranges to be joined
- rest (value|range, required): Value, reference, or range.

Input: `=TEXTJOIN(",", A1, "text1", C1:C5)`

Observed output: `"text1,3,6,9,12,15"`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>CONCAT — value returned</summary>

Concatenates a list or range of text strings

- text1 (string, required): Are 1 to 254 text strings or ranges to be joined to a single text string
- rest (value|range, required): Value, reference, or range.

Input: `=CONCAT("text1", C1:C5, C1:C5)`

Observed output: `"text136912153691215"`

Classification: value returned. Elapsed: 10 ms.

</details>

<details>
<summary>TEXTBEFORE — Excel error</summary>

Returns text that's before delimiting characters.

- text (string, required): The text you want to search for the delimiter.
- delimiter (string, required): The character or string to use as a delimiter.
- instanceNum (value|range, optional): The desired occurrence of delimiter. The default is 1. A negative number searches from the end.
- matchMode (value|range, optional): Searches the text for a delimiter match. By default, a case-sensitive match is done.
- matchEnd (value|range, optional): Whether to match the delimiter against the end of text. By default, they're not matched.
- ifNotFound (value|range, optional): Returned if no match is found. By default, #N/A is returned.

Input: `=TEXTBEFORE("text", ",", A1)`

Observed output: `"#N/A"`

Classification: Excel error. Elapsed: 12 ms.

</details>

<details>
<summary>TEXTAFTER — Excel error</summary>

Returns text that's after delimiting characters.

- text (string, required): The text you want to search for the delimiter.
- delimiter (string, required): The character or string to use as a delimiter.
- instanceNum (value|range, optional): The desired occurrence of delimiter. The default is 1. A negative number searches from the end.
- matchMode (value|range, optional): Searches the text for a delimiter match. By default, a case-sensitive match is done.
- matchEnd (value|range, optional): Whether to match the delimiter against the end of text. By default, they're not matched.
- ifNotFound (value|range, optional): Returned if no match is found. By default, #N/A is returned.

Input: `=TEXTAFTER("text", ",", A1)`

Observed output: `"#N/A"`

Classification: Excel error. Elapsed: 13 ms.

</details>

<details>
<summary>TEXTSPLIT — value returned</summary>

Splits text into rows or columns using delimiters.

- text (string, required): The text to split
- colDelimiter (string, required): Character or string to split columns by.
- rowDelimiter (string, optional): Character or string to split rows by.
- ignoreEmpty (value|range, optional): Whether to ignore empty cells. Defaults to TRUE.
- matchMode (value|range, optional): Searches the text for a delimiter match. By default, a case-sensitive match is done.
- padWith (value|range, optional): The value to use for padding. Defaults to an empty string.

Input: `=TEXTSPLIT("text", ",", ",")`

Observed output: `"text"`

Classification: value returned. Elapsed: 51 ms.

</details>

<details>
<summary>VALUETOTEXT — Excel error</summary>

Returns a text representation of a value

- value (value|range, required): The value to represent as text
- format (string, optional): The format of the text

Input: `=VALUETOTEXT(A1:A5, "#,##0.00")`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 13 ms.

</details>

<details>
<summary>ARRAYTOTEXT — Excel error</summary>

Returns a text representation of an array

- array (array, required): The array to represent as text
- format (string, optional): The format of the text

Input: `=ARRAYTOTEXT({1,2;3,4}, "#,##0.00")`

Observed output: `"#VALUE!"`

Classification: Excel error. Elapsed: 17 ms.

</details>

<details>
<summary>TRANSLATE — unimplemented text</summary>

Translates a text string from one language into another language using the Microsoft Translation service

- text (string, required): Is the text you want to translate
- sourceLanguage (value|range, optional): Specifies the source language your text is currently in. Expressed as a two-letter language code
- targetLanguage (value|range, optional): Specifies the language you want your text to be translated into. Expressed as a two-letter language code

Input: `=TRANSLATE("text", A1)`

Observed output: `"TRANSLATE is not implemented. text=text, sourceLanguage=1, targetLanguage=undefined"`

Classification: unimplemented text. Elapsed: 11 ms.

</details>

<details>
<summary>DETECTLANGUAGE — unimplemented text</summary>

Detects the language of a text string using the Microsoft Translation service

- text (string, required): Is the text you would like to detect the language for

Input: `=DETECTLANGUAGE("text")`

Observed output: `"DETECTLANGUAGE is not implemented. text=text"`

Classification: unimplemented text. Elapsed: 10 ms.

</details>

<details>
<summary>REGEXTEST — value returned</summary>

Checks whether the input matches the pattern, and returns TRUE or FALSE

- text (string, required): Is the text you are searching within
- pattern (value|range, required): The regular expression to be applied
- caseSensitivity (value|range, optional): Whether the match is case sensitive

Input: `=REGEXTEST("text", A1, A1)`

Observed output: `false`

Classification: value returned. Elapsed: 9 ms.

</details>

<details>
<summary>REGEXREPLACE — value returned</summary>

Returns 'text', with 'replacement' in place of matches with 'pattern'

- text (string, required): The text you are searching within
- pattern (value|range, required): The regular expression to be applied
- replacement (value|range, required): The string that replaces the matching substring in text. Where $n appears in the string, where $n is a whole number, $n is replaced with the nth capture group
- occurrence (value|range, optional): Which occurrence to be replaced. If 0, all occurrences are replaced. A negative value means that number of occurrences from the end of the text.
- caseSensitivity (value|range, optional): Whether the match is case sensitive

Input: `=REGEXREPLACE("text", A1, A1)`

Observed output: `"text"`

Classification: value returned. Elapsed: 12 ms.

</details>

<details>
<summary>REGEXEXTRACT — Excel error</summary>

Extracts substrings of 'text' based on the provided REGEX 'pattern'

- text (string, required): Is the text you are searching within
- pattern (value|range, required): The regular expression to be applied
- returnMode (value|range, optional): Specify which matches to return
- caseSensitivity (value|range, optional): Whether the match is case sensitive

Input: `=REGEXEXTRACT("text", A1, A1)`

Observed output: `"#N/A"`

Classification: Excel error. Elapsed: 7 ms.

</details>


## Independent formula expectations

These 77 checks use purpose-built fixtures, not the catalogue smoke fixture.

| Formula | Expected | Observed | Match |
| --- | --- | --- | --- |
| `=SUM(A1:A5)` | 15 | 15 | true |
| `=PRODUCT(A1:A5)` | 120 | 120 | true |
| `=SUMPRODUCT(A1:A5,B1:B5)` | 550 | 550 | true |
| `=SUMIFS(B1:B5,A1:A5,"&gt;2")` | 120 | 120 | true |
| `=COUNTIFS(A1:A5,"&gt;2")` | 3 | 3 | true |
| `=AVERAGEIFS(B1:B5,A1:A5,"&gt;2")` | 40 | 40 | true |
| `=ROUND(12.345,2)` | 12.35 | 12.35 | true |
| `=MOD(17,5)` | 2 | 2 | true |
| `=POWER(2,8)` | 256 | 256 | true |
| `=SQRT(81)` | 9 | 9 | true |
| `=CEILING.MATH(4.2)` | 5 | 5 | true |
| `=FLOOR.MATH(4.8)` | 4 | 4 | true |
| `=ABS(-4)` | 4 | 4 | true |
| `=AVERAGE(A1:A5)` | 3 | 3 | true |
| `=MEDIAN(A1:A5)` | 3 | 3 | true |
| `=MIN(A1:A5)` | 1 | 1 | true |
| `=MAX(A1:A5)` | 5 | 5 | true |
| `=STDEV.S(A1:A5)` | 1.5811388300841898 | 1.5811388300841898 | true |
| `=VAR.P(A1:A5)` | 2 | 2 | true |
| `=CORREL(A1:A5,B1:B5)` | 1 | 1 | true |
| `=NORM.DIST(0,0,1,TRUE)` | 0.5 | 0.5 | true |
| `=PERCENTILE.INC(A1:A5,0.5)` | 3 | 3 | true |
| `=IF(A1=1,"yes","no")` | "yes" | "yes" | true |
| `=AND(A1=1,A2=2)` | true | true | true |
| `=OR(FALSE,TRUE)` | true | true | true |
| `=IFERROR(1/0,99)` | 99 | 99 | true |
| `=IFNA(NA(),99)` | 99 | 99 | true |
| `=CHOOSE(2,10,20,30)` | 20 | 20 | true |
| `=SWITCH(2,1,10,2,20,0)` | 20 | 20 | true |
| `=_xlfn.IFS(A1=1,10,TRUE,20)` | 10 | 10 | true |
| `=INDEX(B1:B5,3)` | 30 | 30 | true |
| `=MATCH(3,A1:A5,0)` | 3 | 3 | true |
| `=VLOOKUP(3,A1:B5,2,FALSE)` | 30 | 30 | true |
| `=XLOOKUP(3,A1:A5,B1:B5)` | 30 | 30 | true |
| `=_xlfn.XLOOKUP(3,A1:A5,B1:B5)` | 30 | 30 | true |
| `=OFFSET(A1,1,1)` | 20 | "OFFSET is not implemented. reference=1, rows=1, cols=1, height=undefined, width=undefined" | false |
| `=INDIRECT("B3")` | 30 | "INDIRECT is not implemented. refText=B3, a1=undefined" | false |
| `=ADDRESS(3,2)` | "$B$3" | "$B$3" | true |
| `=FORMULATEXT(H1)` | "=A1*2" | "FORMULATEXT is not implemented. reference=2" | false |
| `=ISFORMULA(H1)` | true | "ISFORMULA is not implemented. reference=2" | false |
| `=ISNUMBER(B1)` | true | true | true |
| `=ISBLANK(J1)` | true | true | true |
| `=ISERROR(1/0)` | true | true | true |
| `=TYPE("x")` | 2 | 2 | true |
| `=LEN("hello")` | 5 | 5 | true |
| `=LEFT("hello",2)` | "he" | "he" | true |
| `=MID("hello",2,3)` | "ell" | "ell" | true |
| `=TRIM("  a  b  ")` | "a b" | "a b" | true |
| `=SUBSTITUTE("a-b","-","/")` | "a/b" | "a/b" | true |
| `=_xlfn.TEXTJOIN(", ",TRUE,"a","b")` | "a, b" | "a, b" | true |
| `=TEXT(0.125,"0.0%")` | "12.5%" | "12.5%" | true |
| `=DATE(2026,1,1)` | 46023 | 46023 | true |
| `=YEAR(DATE(2026,1,1))` | 2026 | 2026 | true |
| `=MONTH(DATE(2026,9,15))` | 9 | 9 | true |
| `=DAY(EOMONTH(DATE(2026,2,2),0))` | 28 | 28 | true |
| `=MONTH(EDATE(DATE(2026,1,1),2))` | 3 | 3 | true |
| `=NETWORKDAYS(DATE(2026,9,14),DATE(2026,9,18))` | 5 | 5 | true |
| `=DAYS(DATE(2026,2,1),DATE(2026,1,1))` | 31 | 31 | true |
| `=PMT(0,10,1000)` | -100 | -100 | true |
| `=PV(0,10,-100)` | 1000 | 1000 | true |
| `=FV(0,10,-100)` | 1000 | 1000 | true |
| `=NPV(0.1,100,100)` | 173.55371900826447 | 173.55371900826447 | true |
| `=IRR({-100,110})` | 0.1 | 0.1 | true |
| `=XIRR({-100,110},{DATE(2025,1,1),DATE(2026,1,1)})` | 0.1 | 0.09999999999999985 | true |
| `=RATE(1,0,-100,110)` | 0.1 | 0.10000000000000009 | true |
| `=NPER(0,-10,100)` | 10 | 10 | true |
| `=SLN(100,10,9)` | 10 | 10 | true |
| `=CONVERT(1,"m","cm")` | 100 | 100 | true |
| `=DEC2BIN(10)` | "1010" | "1010" | true |
| `=BIN2DEC("1010")` | 10 | 10 | true |
| `=IMABS("3+4i")` | 5 | 5 | true |
| `=IMSUM("1+2i","3+4i")` | "4+6i" | "4+6i" | true |
| `=LET(x,2,x*3)` | 6 | 6 | true |
| `=_xlfn.LET(_xlpm.x,2,_xlpm.x*3)` | 6 | 6 | true |
| `=_xlfn.LAMBDA(_xlpm.x,_xlpm.x*3)(2)` | 6 | 6 | true |
| `=COUNTBLANK(J1:J5)` | 5 | 5 | true |
| `=COUNTIF(J1:J5,"")` | 5 | 0 | false |

### Database criteria fixture

Source rows: North/A/10, South/A/20, North/B/30, South/B/40, with Region/Product/Sales headers and Region=North criteria.

| Formula | Expected | Observed |
| --- | --- | --- | --- |
| `=DSUM(A1:C5,"Sales",E1:E2)` | 40 | 0 |
| `=DCOUNT(A1:C5,"Sales",E1:E2)` | 2 | 4 |
| `=DCOUNTA(A1:C5,"Sales",E1:E2)` | 2 | 4 |
| `=DAVERAGE(A1:C5,"Sales",E1:E2)` | 20 | 25 |
| `=DMIN(A1:C5,"Sales",E1:E2)` | 10 | 10 |
| `=DMAX(A1:C5,"Sales",E1:E2)` | 30 | 40 |
| `=DPRODUCT(A1:C5,"Sales",E1:E2)` | 300 | 240000 |
| `=DSTDEV(A1:C5,"Sales",E1:E2)` | 14.142135623730951 | 12.909944487358056 |
| `=DVAR(A1:C5,"Sales",E1:E2)` | 200 | 166.66666666666666 |
| `=DSTDEVP(A1:C5,"Sales",E1:E2)` | 10 | 11.180339887498949 |
| `=DVARP(A1:C5,"Sales",E1:E2)` | 100 | 125 |
| `=DGET(A1:C5,"Sales",E1:E2)` | #NUM! | #NUM! |

## Detailed probe ledger

Raw “pass” only indicates completion without an exception. A completed discrepancy probe can therefore have that status. Failed initial calls remain visible beside later recovery probes.

### core

| ID | Probe | Harness status | Detail |
| --- | --- | --- | --- |
| C01 | Workbook create, aliases, worksheet collection lifecycle | pass | {"count":2,"index":1,"first":"Beta","nameAt0":"Beta","id":"xdo4zf","sheetId":"1","tabColor":{}} |
| C02 | Worksheet duplicate and missing lookups | pass | {"sameId":true,"count":1,"missing":{}} |
| C03 | Typed matrices, one-dimensional rows/columns, dates, literal formula text | pass | {"values":[[12,true,null],["2026-01-01T00:00:00.000Z",5,"=2+3"]],"formulas":[["","",""],["","=2+3",""]],"rows":[[1,2,3],[4,null,null],[5,null,null],[6,null,null]]} |
| C04 | Broadcast values versus formula assignment | pass | {"formulas":[["=1+1",""],["",""]],"values":[[2,null],[null,null]]} |
| C05 | Single-cell expansion and write resizing | pass | {"rejected":"Error: Range.write would overwrite existing values in Data!A1. Pass overwrite: 'allow' to proceed.","expanded":[[5,6],[7,8]]} |
| C06 | Write clear and resize:none boundary | pass | {"values":[[9,2],[3,4]]} |
| C07 | Index/address/navigation conventions | pass | {"address":"C3:E5","getAddress":"C3:E5","rowIndex":2,"columnIndex":2,"rowCount":3,"columnCount":3,"relativeA1":"A1","byIndex":"D4","offset":"E4:G6","offsetAlias":"E4:G6","resize":"C3:D4","resized":"C3:F6","resizeAlias":"C3:D4","current":"C3:E5","used":"C3:E5"} |
| C08 | Fill down/right shifts relative and preserves anchored references | pass | {"down":[[6],[9],[12]],"right":[["=A1*2","=B1*2","=C1*2"]]} |
| C09 | R1C1, copy all/values/formulas, fillFrom | fail | "Error: Range.fillFrom requires destination to extend source in exactly one direction while sharing the other axis. Source is C1:C2, destination is G1:G2." |
| C10 | Clear modes retain appropriate contents/styles; null removes values | pass | {"values":[[null,null],[null,null]]} |
| C11 | Merge, across merge, worksheet merge aliases | pass | {} |
| C12 | Formatting, theme, sizing, autofit and format matrices | pass | {"rowHeight":22.5,"columnWidth":14.44,"numberFormat":"#,##0.00;(#,##0.00)","theme":{"name":"ChatGPT","colorScheme":{"name":"Probe","colors":[{"name":"accent1","color":{"type":1,"value":"123456"}},{"name":"accent2","color":{"type":1,"value":"007755"}},{"name":"accent3","color":{"type":1,"value":"196B24","transform":{}}},{"name":"accent4","color":{"type":1,"value":"0F9ED5","transform":{}}},{"name":"accent5","color":{"type":1,"value":"A02B93","transform":{}}},{"name":"accent6","color":{"type":1,"value":"4EA72E","transform":{}}},{"name":"dk1","color":{"type":3,"value":"windowText","lastColor":"000000","transform":{}}},{"name":"lt1","color":{"type":3,"value":"window","lastColor":"FFFFFF","transform":{}}},{"name":"dk2","color":{"type":1,"value":"0E2841","transform":{}}},{"name":"lt2","color":{"type":1,"value":"E8E8E8","transform":{}}},{"name":"hlink","color":{"type":1,"value":"467886","transform":{}}},{"name":"folHlink","color":{"type":1,"value":"96607D","transform":{}}},{"name":"bg1","color":{"type":1,"value":"FFFFFF"}},{"name":"tx1","color":{"type":1,"value":"000000"}}]},"backgroundFillStyleList":[{"type":1,"color":{"type":2,"value":"phClr","transform":{}},"gradientStops":[],"pictureEffects":[]},{"type":1,"color":{"type":2,"value":"phClr","transform":{"tint":95000,"satMod":170000}},"gradientStops":[],"pictureEffects":[]},{"type":2,"gradientStops":[{"position":0,"color":{"type":2,"value":"phClr","transform":{"tint":93000,"shade":98000,"lumMod":102000,"satMod":150000}}},{"position":50000,"color":{"type":2,"value":"phClr","transform":{"tint":98000,"shade":90000,"lumMod":103000,"satMod":130000}}},{"position":100000,"color":{"type":2,"value":"phClr","transform":{"shade":63000,"satMod":120000}}}],"pictureEffects":[],"gradientKind":1,"angleDeg":90,"isScaled":false}],"fillStyleList":[{"type":1,"color":{"type":2,"value":"phClr","transform":{}},"gradientStops":[],"pictureEffects":[]},{"type":2,"gradientStops":[{"position":0,"color":{"type":2,"value":"phClr","transform":{"lumMod":110000,"satMod":105000,"tint":67000}}},{"position":50000,"color":{"type":2,"value":"phClr","transform":{"lumMod":105000,"satMod":103000,"tint":73000}}},{"position":100000,"color":{"type":2,"value":"phClr","transform":{"lumMod":105000,"satMod":109000,"tint":81000}}}],"pictureEffects":[],"gradientKind":1,"angleDeg":90,"isScaled":false},{"type":2,"gradientStops":[{"position":0,"color":{"type":2,"value":"phClr","transform":{"satMod":103000,"lumMod":102000,"tint":94000}}},{"position":50000,"color":{"type":2,"value":"phClr","transform":{"satMod":110000,"lumMod":100000,"shade":100000}}},{"position":100000,"color":{"type":2,"value":"phClr","transform":{"lumMod":99000,"satMod":120000,"shade":78000}}}],"pictureEffects":[],"gradientKind":1,"angleDeg":90,"isScaled":false}],"lineStyleList":[{"style":1,"widthEmu":12700,"fill":{"type":1,"color":{"type":2,"value":"phClr","transform":{}},"gradientStops":[],"pictureEffects":[]}},{"style":1,"widthEmu":19050,"fill":{"type":1,"color":{"type":2,"value":"phClr","transform":{}},"gradientStops":[],"pictureEffects":[]}},{"style":1,"widthEmu":25400,"fill":{"type":1,"color":{"type":2,"value":"phClr","transform":{}},"gradientStops":[],"pictureEffects":[]}}],"effectStyleList":[{"effects":[]},{"effects":[]},{"effects":[{"type":1,"shadow":{"color":{"type":1,"value":"000000","transform":{"alpha":63000}},"blurRadius":57150,"distance":19050,"direction":5400000}}]}]}} |
| C13 | Theme/RGB transformed colors, gradient fills, borders and alignments | pass | {"accepted":true} |
| C14 | Freeze rows and columns, unfreeze, worksheet reset | pass | {"reset":true} |
| C15 | Named range calculation, scopes and deletion | pass | {"address":"A1:A3"} |
| C16 | Named LAMBDA functions and prefixed variables | pass | {"value":"#NAME?","names":"{\"kind\":\"definedName\",\"name\":\"AddTax\",\"scope\":\"workbook\"}"} |
| C17 | Threaded cell/range comments, reply, reactions and state | pass | {"inspect":"{\"kind\":\"thread\",\"id\":\"th/{395A0A20-231A-4828-9E91-D8E5C3D9E552}\",\"sheet\":\"Data\",\"target\":\"A1\",\"status\":\"1\",\"text\":\"Synthetic review\",\"comments\":[{\"id\":\"{395A0A20-231A-4828-9E91-D8E5C3D9E552}\",\"text\":\"Synthetic review\",\"authorId\":\"{1B2931D3-F392-4E97-BB3D-87680D956BAE}\",\"createdAt\":\"2026-09-15T17:15:14.635Z\"},{\"id\":\"{312D6596-1E4F-459E-B41E-58AD9AB19E5E}\",\"text\":\"Checked\",\"authorId\":\"{1B2931D3-F392-4E97-BB3D-87680D956BAE}\",\"createdAt\":\"2026-09-15T17:15:14.635Z\"}]}\n{\"kind\":\"thread\",\"id\":\"th/{36BD5F07-B3F3-4056-9618-506AAF275C4A}\",\"sheet\":\"Data\",\"target\":\"A2:B2\",\"status\":\"1\",\"text\":\"Range review\",\"comments\":[{\"id\":\"{36BD5F07-B3F3-4056-9618-506AAF275C4A}\",\"text\":\"Range review\",\"authorId\":\"{1B2931D3-F392-4E97-BB3D-87680D956BAE}\",\"createdAt\":\"2026-09-15T17:15:14.635Z\"}]}"} |
| C18 | Tables create, append, style, header/totals and delete | fail | "AssertionError [ERR_ASSERTION]: 5 != 9" |
| C19 | Headerless table and overlap acceptance boundary | pass | {"headerless":[["A",1],["B",3],["C",5]],"overlap":"accepted"} |
| C20 | CSV quote/newline parsing and string typing | pass | {"before":[["Name","Qty","Note"],["North, Ltd","2","Line 1\nLine 2"],["South","03","plain"]],"converted":5,"append":"accepted"} |
| C21 | Instance CSV and Markdown import | pass | {"instanceRange":"A1:B2","instanceSheet":"CSV","markdown":[["Name","Value"],["Alpha","10"],["Beta","20"]],"tableCount":1} |
| C22 | HTML copy/paste values versus formulas | pass | {"result":{"range":"C3:D4","selectionRect":{"r1":2,"c1":2,"r2":3,"c2":3}},"values":[["Qty","Double"],[5,0]],"formulas":[["",""],["","=A2*2"]],"htmlLength":1287} |
| C23 | Serialize/load, validated load, constructor and utilities | pass | {"protoKeys":["id","sheets","styles","featurePropertyBags","theme","contentReferences","images","people","threads","notes","slicerCaches","pivotCaches","timelineCaches","metadata","definedNames","textStyles","codeEnvironments","codeBlocks"],"address":"A1:B3","fillRight":[[1,1,1]],"fillDown":[[1],[1],[1]]} |
| C24 | Cross-sheet dependencies, recalculation, trace and stats | pass | {"trace":{"cell":"Build!A1","sheetName":"Build","address":"A1","formula":"='Input'!B1*(1+'Input'!B2)","value":120,"params":[{"cell":"Input!B1","sheetName":"Input","address":"B1","formula":null,"value":100,"params":[]},{"cell":"Input!B2","sheetName":"Input","address":"B2","formula":null,"value":0.2,"params":[]}]},"stats":{},"badTrace":null} |
| C25 | Dynamic arrays, spill descriptors and obstruction | pass | {"before":{"formulas":[["=SEQUENCE(3,2)",""],["",""],["",""]],"display":[["=SEQUENCE(3,2)","=SEQUENCE(3,2)"],["=SEQUENCE(3,2)","=SEQUENCE(3,2)"],["=SEQUENCE(3,2)","=SEQUENCE(3,2)"]],"infos":[[{"kind":"stored","formula":"=SEQUENCE(3,2)","display":"=SEQUENCE(3,2)","isEditable":true},{"kind":"projected","source":"spill","display":"=SEQUENCE(3,2)","anchor":"Data!A1","ref":"Data!A1:B3","isEditable":false}],[{"kind":"projected","source":"spill","display":"=SEQUENCE(3,2)","anchor":"Data!A1","ref":"Data!A1:B3","isEditable":false},{"kind":"projected","source":"spill","display":"=SEQUENCE(3,2)","anchor":"Data!A1","ref":"Data!A1:B3","isEditable":false}],[{"kind":"projected","source":"spill","display":"=SEQUENCE(3,2)","anchor":"Data!A1","ref":"Data!A1:B3","isEditable":false},{"kind":"projected","source":"spill","display":"=SEQUENCE(3,2)","anchor":"Data!A1","ref":"Data!A1:B3","isEditable":false}]]},"childEdit":"accepted","blocked":[[1,2],[3,4],[5,6]]} |
| C26 | Inspection kind families, bounding, regex error search and resolve | pass | {"sheet":"{\"kind\":\"sheet\",\"id\":\"ws/1gz30i\",\"name\":\"Data\",\"index\":0,\"range\":\"A1:B3\",\"address\":\"A1:B3\"}","find":{"matches":[{"kind":"match","sheet":"Data","address":"A2","value":"East","formula":null,"match":"value"}],"truncated":false,"offset":0,"limit":2,"total":1,"notices":[]},"errorMatch":"{\"kind\":\"match\",\"sheet\":\"Data\",\"address\":\"B3\",\"value\":\"#DIV/0!\",\"formula\":\"1/0\",\"match\":\"value\"}"} |
| C27 | Error scan versus missing-data and full-column boundaries | pass | [{"formula":"=COUNTIF(A1:A4,\"\")","value":0},{"formula":"=COUNTIFS(A1:A4,\"\")","value":0},{"formula":"=COUNTBLANK(A1:A4)","value":1},{"formula":"=SUMIFS(B:B,A:A,\"x\")","value":"#VALUE!"},{"formula":"=SUMIFS(B1:B4,A1:A4,\"x\")","value":3},{"formula":"=ROWS(A:A)","value":4},{"formula":"=1/0","value":"#DIV/0!"},{"formula":"=NA()","value":"#N/A"},{"formula":"=MISSINGFUNCTION(1)","value":"#NAME?"}] |
| C28 | Recorded patch/replay and patch input envelope | pass | {"result":"done","patch":[{"op":"range.values.set","target":{"sheet":"Data","range":"A1"},"values":[[2]]},{"op":"range.formulas.set","target":{"sheet":"Data","range":"B1"},"formulas":[["=A1*3"]]}],"idMap":{},"crdtBytes":104,"applied":{"idMap":{},"warnings":[]}} |
| C29 | Collaborative state, CRDT subscription and replica application | pass | {"ready":true,"updates":[{"bytes":45,"origin":"[object Object]"},{"bytes":63,"origin":"[object Object]"}],"recordedBytes":79,"replicaValue":1} |
| C30 | Native export and import preserve values, formulas and update behavior | pass | {"formulas":[["=A2*B2"],["=A3*B3"]],"values":[[40],[30]]} |
### features-first

| ID | Probe | Harness status | Detail |
| --- | --- | --- | --- |
| F01 | All conditional-format rule families; serialize, render and clear | pass | [{"type":"cellIs","accepted":true},{"type":"CellValue","accepted":true},{"type":"Custom","accepted":true},{"type":"expression","accepted":true},{"type":"colorScale","accepted":true},{"type":"dataBar","accepted":true},{"type":"iconSet","accepted":true},{"type":"containsText","accepted":true},{"type":"notContainsText","accepted":true},{"type":"beginsWith","accepted":true},{"type":"endsWith","accepted":true},{"type":"containsBlanks","accepted":true},{"type":"notContainsBlanks","accepted":true},{"type":"containsErrors","accepted":true},{"type":"notContainsErrors","accepted":true},{"type":"duplicateValues","accepted":true},{"type":"uniqueValues","accepted":true},{"type":"timePeriod","accepted":true},{"type":"top10","accepted":true},{"type":"aboveAverage","accepted":true}] |
| F02 | All validation types, settings, collection and Office-style assignment | pass | {"types":["none","whole","decimal","list","date","time","textLength","custom"],"invalidProgrammaticWriteAccepted":true} |
| F03 | All sparkline types, range alias, markers, axes and export | fail | "TypeError: extra.delete is not a function" |
| F04 | Data table full-range convention and independent two-variable results | pass | {"resultType":"undefined","actual":[[20,30,40],[30,40,50],[40,50,60]],"formulas":[["","",""],["","",""],["","",""]],"display":[["{=TABLE(B2,B3)}","{=TABLE(B2,B3)}","{=TABLE(B2,B3)}"],["{=TABLE(B2,B3)}","{=TABLE(B2,B3)}","{=TABLE(B2,B3)}"],["{=TABLE(B2,B3)}","{=TABLE(B2,B3)}","{=TABLE(B2,B3)}"]],"infos":[[{"kind":"projected","source":"dataTable","display":"{=TABLE(B2,B3)}","master":"Sensitivity!F5","ref":"Sensitivity!F5:H7","rowInput":"B2","columnInput":"B3","isEditable":false}]]} |
| F05 | Data table older body-only and one-variable documentation | pass | {"body":"Error: Data table range needs a top-left formula cell. \"F5\" is empty. Did you mean \"E4:H7\"?","one":[[20,1,2,3],[null,10,20,30]]} |
| F06 | Drawing images by SVG, bytes, data URL, path; replace and delete | fail | "TypeError: this.anchor.toProtoFields is not a function" |
| F07 | Shapes, text, drawing layout directions and deleteAllDrawings | pass | [{"direction":"vertical","count":4},{"direction":"horizontal","count":4},{"direction":"grid","count":4}] |
| CH-line | Chart type acceptance, export and render: line | pass | {"type":"line","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":18130}} |
| CH-pie | Chart type acceptance, export and render: pie | pass | {"type":"pie","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":13648}} |
| CH-bar | Chart type acceptance, export and render: bar | pass | {"type":"bar","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":15863}} |
| CH-doughnut | Chart type acceptance, export and render: doughnut | pass | {"type":"doughnut","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":15545}} |
| CH-scatter | Chart type acceptance, export and render: scatter | pass | {"type":"scatter","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":14132}} |
| CH-bubble | Chart type acceptance, export and render: bubble | pass | {"type":"bubble","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":14471}} |
| CH-radar | Chart type acceptance, export and render: radar | pass | {"type":"radar","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":19975}} |
| CH-treemap | Chart type acceptance, export and render: treemap | pass | {"type":"treemap","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":15433}} |
| CH-sunburst | Chart type acceptance, export and render: sunburst | pass | {"type":"sunburst","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":18345}} |
| CH-map | Chart type acceptance, export and render: map | pass | {"type":"map","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":9867}} |
| CH-waterfall | Chart type acceptance, export and render: waterfall | pass | {"type":"waterfall","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":17399}} |
| CH-line3D | Chart type acceptance, export and render: line3D | pass | {"type":"line3D","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":19044}} |
| CH-pie3D | Chart type acceptance, export and render: pie3D | pass | {"type":"pie3D","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":14232}} |
| CH-area3D | Chart type acceptance, export and render: area3D | pass | {"type":"area3D","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":15870}} |
| CH-bar3D | Chart type acceptance, export and render: bar3D | pass | {"type":"bar3D","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":10414}} |
| CH-funnel | Chart type acceptance, export and render: funnel | pass | {"type":"funnel","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":12560}} |
| CH-histogram | Chart type acceptance, export and render: histogram | pass | {"type":"histogram","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":12812}} |
### features-remaining

| ID | Probe | Harness status | Detail |
| --- | --- | --- | --- |
| CH-stock | Chart type acceptance, export and render: stock | pass | {"type":"stock","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":16716}} |
| CH-surface3D | Chart type acceptance, export and render: surface3D | pass | {"type":"surface3D","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":17544}} |
| CH-ofPie | Chart type acceptance, export and render: ofPie | pass | {"type":"ofPie","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":16530}} |
| CH-surface | Chart type acceptance, export and render: surface | pass | {"type":"surface","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":17011}} |
| CH-pareto | Chart type acceptance, export and render: pareto | pass | {"type":"pareto","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":16859}} |
| CH-combo | Chart type acceptance, export and render: combo | pass | {"type":"combo","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":16574}} |
| CH-area | Chart type acceptance, export and render: area | pass | {"type":"area","series":[{"name":"Series A","formula":"'Chart'!$B$2:$B$5","categoryFormula":"'Chart'!$A$2:$A$5"},{"name":"Series B","formula":"'Chart'!$C$2:$C$5","categoryFormula":"'Chart'!$A$2:$A$5"}],"render":{"bytes":15091}} |
| F08 | Charts range/nonadjacent/row binding, setData, axes, labels, legend, line and data table | fail | "TypeError: this[#R].toProto is not a function" |
| F09 | Raster export selection and output variants | pass | [{"name":"png-index","type":"image/png","size":11938},{"name":"jpeg-center","type":"image/jpeg","size":9977},{"name":"layout","type":"application/vnd.openai.workbook-layout+json","size":3884},{"name":"xlsx","type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","size":3432}] |
### followup

| ID | Probe | Harness status | Detail |
| --- | --- | --- | --- |
| R01 | FillFrom correct adjacent extension and table totals boundary | pass | {"fill":[["=A1+B1"],["=A2+B2"],["=A3+B3"]],"before":{"range":"E1:F4","rows":[["A",2],["B",3],["C",4]]},"after":{"range":"E1:F4","rows":[["A",2],["B",3],["C",4]],"grid":[["Item","Qty"],["A",2],["B",3],["C",4]]},"insertError":"Error: Inserting table rows at a specific index is not supported yet","columnsAPI":"undefined","filterAPI":"undefined"} |
| R02 | Named function formula-versus-lambda configuration | pass | [{"name":"Plain","config":{"lambda":"LAMBDA(amount,amount*1.1)"},"value":110.00000000000001},{"name":"Prefixed","config":{"lambda":"_xlfn.LAMBDA(_xlpm.amount,_xlpm.amount*1.1)"},"value":110.00000000000001},{"name":"Equals","config":{"lambda":"=_xlfn.LAMBDA(_xlpm.amount,_xlpm.amount*1.1)"},"value":110.00000000000001}] |
| R03 | CSV instance import on populated workbook preserves or replaces data | fail | "Error: hydrateCrdtFromProto requires an empty collaborative document." |
| R04 | One-variable row and column data tables; export and roundtrip | pass | {"row":[[10,20,30]],"column":[[10],[20],[30]],"reimport":[[20,1,2,3],[null,10,20,30]]} |
| R05 | Sparklines supported delete route and export preservation | pass | {"count":3} |
| R06 | Image sources and config replacement without overwriting facade anchor | pass | {"count":5,"fromImage":{"range":"F1","selectionRect":{"r1":0,"c1":5,"r2":0,"c2":5},"imageId":"e1efc9ba-80e5-45e0-9c31-b52a6e1732d5","widthPx":590,"heightPx":220},"types":["image/svg+xml","image/png","image/png","image/svg+xml","image/png"]} |
| R07 | Chart configuration at creation, mutable facade settings and analytics | fail | "AssertionError [ERR_ASSERTION]: Expected values to be strictly deep-equal:\n\nundefined !== 200\n" |
| R08 | Chart dataLabels and dataTable config-first API | pass | {"labels":true,"dataTable":false} |
| R09 | Pivots and slicers, hierarchy operations and native export | fail | "Error: PivotTable SalesPivot not found" |
| R10 | Notes operation route and low-level note persistence | pass | {"exported":true} |
| R11 | Undo/redo, recordAsync and awareness selections | pass | {"result":7,"patch":[{"op":"range.values.set","target":{"sheet":"Data","range":"A1"},"values":[[2]]},{"op":"range.formulas.set","target":{"sheet":"Data","range":"B1"},"formulas":[["=A1*2"]]}],"before":{"canUndo":true,"canRedo":false,"value":2},"undone":{"canRedo":true,"value":1},"redone":2,"selections":[{"presenceId":"synthetic-reviewer","kind":"collaborator","sheetName":"Data","rect":{},"selectedElementIds":[]}]} |
### advanced

| ID | Probe | Harness status | Detail |
| --- | --- | --- | --- |
| A01 | Pivot root collection, native serialization and slicer fields | pass | {"grid":[["Row Labels","","","Grand Total",null],["North",10,30,40,null],["South",20,40,60,null],["Grand Total",30,70,100,null],[null,null,null,null,null],[null,null,null,null,null],[null,null,null,null,null]],"hierarchies":["Region","Product","Sales"],"methods":["constructor","workbook","worksheet","name","layout","hierarchies","rowHierarchies","columnHierarchies","dataHierarchies","filterHierarchies","_assignHierarchy","_getHierarchies","_getHierarchyByFieldIndex","_ensureHierarchyIndex","delete","cache","toProto","rebuildCache","addDeleteListener","removeDeleteListener","addNameChangeListener","removeNameChangeListener","__getLayoutType","__setLayoutType","__getLayoutMeta","__setLayoutMeta","__getLocationForLayout","__getDataBodyShapeForLayout","__setRowItems","__setColumnItems","__clearPivotCellRenderHints","__setPivotCellRenderHint","__getPivotCellRenderHint","__getPivotCellRenderHints","__findDataFieldProto","__addDataFieldProto","__getPivotFieldProtos","__updateLocation","__applyFilterForField","__clearAllFiltersForField","preview"],"pivotExport":"completed","slicerMethods":["constructor","name","caption","left","top","height","width","style","selectItems","clearFilters","delete","addDeleteListener","removeDeleteListener","addNameChangeListener","removeNameChangeListener"],"slicerExport":"completed"} |
| A02 | Chart analytics and mutable label/table objects | pass | {"resolvedValues":[],"formula":"'Charts'!$B$2:$B$5","dataTable":true,"chartCount":1,"converted":{"bytes":{},"contentType":"image/png","widthPx":590,"heightPx":360},"chartCountAfter":1,"imageCountAfter":0} |
| A03 | Native note through operation API and note delete | fail | "TypeError: Cannot read properties of undefined (reading 'proto')" |
| A04 | Awareness operation families and correct presence rectangle | pass | {"result":{"idMap":{},"warnings":["TODO: presence.selection.set is not implemented yet."]},"selected":[],"reaction":{"idMap":{},"warnings":["TODO: thread.reaction.toggle is not implemented yet."]}} |
| A05 | Workbook static session API and common CRDT initialization | pass | {"statics":[],"methods":{},"crdt":{"loadInitial":"loadInitialCrdtStateV2(t,n={}){this.#lt(\"loadInitialCrdtStateV2\");this.applyCrdtUpdateV2(t,n)}","subscribe":"onCrdtUpdateV2(t){const n=(r,i)=&gt;{t(new Uint8Array(r),i)};this.#I.on(\"updateV2\",n);return()=&gt;{this.#I.off(\"updateV2\",n)}}","getDoc":"getCrdtDoc(){return this.#I}"}} |
| A06 | Database formula family with valid database and criteria | pass | [{"formula":"=DSUM(A1:C5,\"Sales\",E1:E2)","expected":40,"actual":0},{"formula":"=DCOUNT(A1:C5,\"Sales\",E1:E2)","expected":2,"actual":4},{"formula":"=DCOUNTA(A1:C5,\"Sales\",E1:E2)","expected":2,"actual":4},{"formula":"=DAVERAGE(A1:C5,\"Sales\",E1:E2)","expected":20,"actual":25},{"formula":"=DMIN(A1:C5,\"Sales\",E1:E2)","expected":10,"actual":10},{"formula":"=DMAX(A1:C5,\"Sales\",E1:E2)","expected":30,"actual":40},{"formula":"=DPRODUCT(A1:C5,\"Sales\",E1:E2)","expected":300,"actual":240000},{"formula":"=DSTDEV(A1:C5,\"Sales\",E1:E2)","expected":14.142135623730951,"actual":12.909944487358056},{"formula":"=DVAR(A1:C5,\"Sales\",E1:E2)","expected":200,"actual":166.66666666666666},{"formula":"=DSTDEVP(A1:C5,\"Sales\",E1:E2)","expected":10,"actual":11.180339887498949},{"formula":"=DVARP(A1:C5,\"Sales\",E1:E2)","expected":100,"actual":125},{"formula":"=DGET(A1:C5,\"Sales\",E1:E2)","expected":"#NUM!","actual":"#NUM!"}] |
| A07 | Table total reserve and recalculation before/after flag | pass | {"before":9,"after":9,"withReservedTotalsRow":9} |
### recovery

| ID | Probe | Harness status | Detail |
| --- | --- | --- | --- |
| V01 | Replicated CRDT edits from a shared initialization update | pass | {"baselineUpdates":[5761],"sheets":["Data"],"value":99} |
| V02 | Correct raw-note config and Excel note preservation | pass | {"id":"synthetic-note","target":{"cell":{"sheetName":"Data","sheetId":"1","address":"A1"}},"authorId":"{069BF49E-8F04-4036-917B-00633CDA01A2}","createdAt":"2026-09-15T17:28:16.054Z","body":{"plainText":"Synthetic cell note"}} |
| V03 | Unsupported operations return warnings; batches retain earlier changes | pass | {"result":{"idMap":{},"warnings":["TODO: invalid.action is not implemented yet.","TODO: note.add is not implemented yet."]},"value":2,"notes":0} |
| V04 | Fresh structured formula after table totals flag and saved-file reload | pass | {"before":9,"after":[[9],[5]],"reimport":[[5],[5]]} |
| V05 | Slicer select/clear, pivot refresh, layout and export | pass | {"before":[["Row Labels","Sum of Sales",null],["North",40,null],["South",60,null],["Grand Total",100,null],[null,null,null]],"selected":[["Row Labels","Sum of Sales",null],["North",40,null],["Grand Total",40,null],["Grand Total",100,null],[null,null,null]],"refreshed":[["Row Labels","Sum of Sales",null],["North",130,null],["South",60,null],["Grand Total",190,null],[null,null,null]],"layout":"Compact"} |
| V06 | Read-only public static session wrapper types and file APIs | pass | {"Workbook":{"run":{"type":"function","source":"run(t,n,r){return e.session.run(t,n,r)}"},"create":{"type":"function","source":"create(){if(arguments.length&gt;0){throw new Error(\"Workbook.create() does not accept worksheet names; add worksheets with workbook.worksheets.add(name).\")}return new e({sheets:[],styles:void 0,theme:void 0,contentReferences:[],images:[],people:[],threads:[],notes:[],slicerCaches:[],pivotCaches:[],timelineCaches:[],textStyles:[],codeEnvironments:[],codeBlocks:[]})}"},"load":{"type":"function","source":"load(t,n){const r=_So(n);if(r){const i=pfr(t,r);if(!i.valid){throw new Bae(i)}}return new e(t)}"},"fromGoogleSheets":{"type":"function","source":"async fromGoogleSheets(t){return gze().loadWorkbook(t)}"},"fromMarkdown":{"type":"function","source":"async fromMarkdown(t,n){const{buildTableValuesFromMarkdown:r,sanitizeSheetName:i}=await a0r();const o=r(t);const a=n?.format??true;const s=e.create();const l=s.worksheets.add(i(n?.sheetName));const u=o.reduce((d,f)=&gt;Math.max(d,f.length),0);if(u===0){return s}for(let d=0;d&lt;o.length;d+=1){const f=o[d]??[];const h=f.slice();while(h.length&lt;u){h.push(\"\")}const m=`${Cf(0)}${d+1}`;const g=`${Cf(u-1)}${d+1}`;const b=u===1?m:`${m}:${g}`;l.getRange(b).values=[h]}if(o.length&gt;0&&u&gt;0){const d=`${Cf(0)}1`;const f=`${Cf(u-1)}${o.length}`;const h=u===1&&o.length===1?d:`${d}:${f}`;if(a){l.tables.add(h,true)}l.getRange(h).format.autofitColumns()}return s}"},"fromCSV":{"type":"function","source":"async fromCSV(t,n){const{planCsvImport:r}=await f_t();const i=r(t,n);const o=e.create();const a=o.worksheets.add(i.sheetName);if(i.values.length&gt;0&&i.rangeRef){const s=i.rangeRef;o.record(()=&gt;{a.getRange(s).values=i.values})}return o}"},"session":{"type":"object"}},"SpreadsheetFile":["length","name","prototype","importXlsx","exportXlsx"],"FileBlob":["length","name","prototype","load"]} |

## Runtime facade inventory

This inventory identifies methods/getters visible on the installed facade prototypes. It is discovery evidence, not a promise that undocumented members are stable or that every method has been independently verified. Private-prefixed helpers are excluded here. Public feature examples, operations and corrective probes are the usage evidence.

### workbook

`session` (getter), `attachSessionController` (method), `attachArtifactSessionOutputSink` (method), `useExternalImageAssetStorage` (method), `fromCSV` (method), `pivotTables` (getter), `fontFamilies` (getter), `sheets` (getter), `slicers` (getter), `worksheets` (getter), `getStyleRegistry` (method), `getSpreadsheetRenderAssets` (method), `hydrateCrdtFromProto` (method), `loadInitialCrdtStateV2` (method), `getCrdtDoc` (method), `isCollaborativeStateReady` (method), `getCollabOrigins` (method), `queueWorkbookCollabPublish` (method), `queueWorksheetCollabPublish` (method), `runLocalCollabTransaction` (method), `batchCellInputWrites` (method), `undo` (method), `redo` (method), `canUndo` (method), `canRedo` (method), `getConditionalFormattingRenderCache` (method), `invalidateConditionalFormattingCache` (method), `createImageAsset` (method), `createImageAssetModel` (method), `getImageAssetItems` (method), `getPendingImageHydrationRequests` (method), `hydrateImageAssets` (method), `images` (getter), `resolveImageAsset` (method), `notes` (getter), `names` (getter), `definedNames` (getter), `utils` (getter), `theme` (getter), `setColorScheme` (method), `recalculate` (method), `trace` (method), `collectFormulaUsageStats` (method), `toProto` (method), `inspect` (method), `findCells` (method), `help` (method), `resolve` (method), `apply` (method), `configureGoogleSheets` (method), `applyCrdtUpdateV2` (method), `onCrdtUpdateV2` (method), `record` (method), `recordAsync` (method), `getRecorder` (method), `googleSheets` (getter), `render` (method), `export` (method), `toHTML` (method), `fromHTML` (method), `chartToImage` (method), `fromImage` (method)

### worksheets

`getSheetCount` (method), `getSheetIndex` (method), `getSheetNameByIndex` (method), `getItemAt` (method), `getFirst` (method), `getOrAdd` (method), `add` (method), `getItem` (method), `getItemOrNullObject` (method), `getItemBySheetId` (method), `getTopologyEntries` (method), `syncTopology` (method), `items` (getter), `getActiveWorksheet` (method), `getActive` (method), `setActiveWorksheet` (method), `toProto` (method)

### worksheet

`workbook` (getter), `names` (getter), `id` (getter), `name` (getter, writable), `isNullObject` (getter), `index` (getter, writable), `sheetId` (getter), `tabColor` (getter, writable), `showGridLines` (getter, writable), `defaultRowHeight` (getter), `defaultColWidth` (getter), `baseColWidth` (getter), `innerXml` (getter), `outerXml` (getter), `writeName` (method), `writeIndex` (method), `getRange` (method), `getRangeByIndexes` (method), `getCell` (method), `getUsedRange` (method), `reset` (method), `mergeCells` (method), `unmergeCells` (method), `pivotTables` (getter), `slicers` (getter), `sparklineGroups` (getter), `charts` (getter), `shapes` (getter), `images` (getter), `deleteAllDrawings` (method), `autoLayoutDrawings` (method), `conditionalFormattings` (getter), `dataValidations` (getter), `freezePanes` (getter), `sparklines` (getter), `tables` (getter), `dataTables` (getter), `cells` (getter), `flushCollaborativeState` (method), `delete` (method), `toProto` (method), `writeCellInputToYjs` (method), `hydrateCellInputToYjs` (method), `hydrateCollaborativeRefsFromProto` (method)

### range

`address` (getter), `getAddress` (method), `write` (method), `rowCount` (getter), `getRowCount` (method), `columnCount` (getter), `getColumnCount` (method), `format` (getter, writable), `rowIndex` (getter), `columnIndex` (getter), `getRowIndex` (method), `getColumnIndex` (method), `getRange` (method), `setNumberFormat` (method), `sparklines` (getter), `conditionalFormats` (getter), `dataValidation` (getter, writable), `getCell` (method), `getRow` (method), `getColumn` (method), `getRangeByIndexes` (method), `getOffsetRange` (method), `getResizedRange` (method), `getResizeRange` (method), `offset` (method), `resize` (method), `getCurrentRegion` (method), `formulas` (getter, writable), `formulasR1C1` (getter, writable), `values` (getter, writable), `displayFormula` (getter), `displayFormulas` (getter), `formulaInfo` (method), `formulaInfos` (getter), `rawValues` (getter), `writeValues` (method), `getLastRow` (method), `load` (method), `calculate` (method), `clear` (method), `merge` (method), `unmerge` (method), `fillDown` (method), `fillRight` (method), `fillFrom` (method), `copyFrom` (method), `copyTo` (method), `getBoundingBox` (method)

### format

`reset` (method), `font` (getter, writable), `setFont` (method), `borders` (getter, writable), `fill` (getter, writable), `numberFormat` (getter, writable), `wrapText` (getter, writable), `horizontalAlignment` (getter, writable), `verticalAlignment` (getter, writable), `rowHeight` (getter, writable), `rowHeightPx` (getter, writable), `columnWidth` (getter, writable), `columnWidthPx` (getter, writable), `styleId` (getter), `autofitColumns` (method), `autofitRows` (method), `applyBorderBlueprint` (method), `updateBorders` (method), `getBorderBlueprintSnapshot` (method)

### tables

`add` (method), `getItem` (method), `getItemOrNullObject` (method), `getItemAt` (method), `items` (getter), `replaceFromProto` (method), `deleteAll` (method)

### table

`isNullObject` (getter), `worksheet` (getter), `hasHeaders` (getter), `showHeaders` (getter, writable), `writeShowHeaders` (method), `name` (getter, writable), `writeName` (method), `address` (getter), `style` (getter, writable), `writeStyle` (method), `showTotals` (getter, writable), `writeShowTotals` (method), `showBandedRows` (getter, writable), `writeShowBandedRows` (method), `highlightFirstColumn` (getter, writable), `writeHighlightFirstColumn` (method), `highlightLastColumn` (getter, writable), `writeHighlightLastColumn` (method), `showBandedColumns` (getter, writable), `writeShowBandedColumns` (method), `showFilterButton` (getter, writable), `writeShowFilterButton` (method), `getRange` (method), `getHeaderRowRange` (method), `rows` (getter), `delete` (method), `appendRows` (method), `getDataRows` (method), `refreshBounds` (method), `syncColumnsFromSheet` (method)

### tableRows

`add` (method), `items` (getter)

### tableColumns

No facade at this property in the tested object.

### notes

`add` (method), `items` (getter), `proto` (getter), `replace` (method)

### comments

`self` (getter), `setSelf` (method), `clearSelf` (method), `addThread` (method), `getThread` (method), `toProto` (method), `replaceFromProto` (method)

### names

`hasOwnProperty` (method), `isPrototypeOf` (method), `propertyIsEnumerable` (method), `toString` (method), `valueOf` (method), `toLocaleString` (method)

### pivotTables

`add` (method), `getItem` (method), `getItemOrNullObject` (method), `items` (getter), `reload` (method), `hydrateLayouts` (method)

### slicers

`add` (method), `getItem` (method), `getItemAt` (method), `items` (getter), `reload` (method), `synthesizeCaches` (method)

### awareness

`getSheetState` (method), `setActiveSheetName` (method), `renameSheet` (method), `forgetSheet` (method), `setDialog` (method), `setPresenceSelections` (method), `clearPresenceSelections` (method), `getSelectionsForSheet` (method), `recordMissingAction` (method), `pushUndoEntry` (method), `popUndoEntry` (method), `pushRedoEntry` (method), `popRedoEntry` (method), `clearHistory` (method)

### charts

`items` (getter), `getItem` (method), `getItemAt` (method), `getItemOrNullObject` (method), `deleteAll` (method), `clear` (method), `add` (method), `toProto` (method), `hydrateFromFallbackState` (method)

### chart

`aid` (getter), `type` (getter, writable), `titleText` (getter, writable), `title` (getter, writable), `categories` (getter, writable), `legend` (getter, writable), `axes` (getter), `plotArea` (getter), `chartArea` (getter), `isNullObject` (getter), `width` (getter, writable), `height` (getter, writable), `setPosition` (method), `setData` (method), `delete` (method), `resolveBoundsPx` (method), `captureAnchorSnapshot` (method), `restoreAnchorSnapshot` (method), `setPreviewBoundsPx` (method), `clearPreviewBounds` (method), `applyBoundsPx` (method), `toDrawingProto` (method)

### series

`add` (method), `setChangeHandler` (method), `getItemAt` (method), `deleteAt` (method), `items` (getter), `clear` (method), `deleteAll` (method), `toProto` (method), `length` (getter)

### seriesItem

`setChangeHandler` (method), `name` (getter, writable), `values` (getter, writable), `resolveValues` (method), `categories` (getter, writable), `resolveCategories` (method), `xValues` (getter, writable), `xFormula` (getter, writable), `categoryPaths` (getter, writable), `marker` (getter, writable), `bubbleSizes` (getter, writable), `valuesFormatCode` (getter, writable), `xValuesFormatCode` (getter, writable), `fill` (getter, writable), `format` (getter), `explosion` (getter, writable), `smooth` (getter, writable), `stroke` (getter, writable), `line` (getter, writable), `setCategoryNormalizer` (method), `dataLabelOverrides` (getter, writable), `dataLabels` (getter, writable), `trendlines` (getter, writable), `errorBars` (getter, writable), `points` (getter, writable), `formula` (getter, writable), `categoryFormula` (getter, writable), `isNullObject` (getter), `delete` (method), `toProto` (method)

### shapes

`items` (getter), `deleteAll` (method), `add` (method), `toProto` (method), `replace` (method)

### images

`items` (getter), `deleteAll` (method), `add` (method), `toProto` (method), `replace` (method)

### sparklines

`add` (method), `getAll` (method), `getGroupForCell` (method), `clear` (method), `deleteAll` (method), `delete` (method), `toProto` (method), `renderContext` (getter), `replaceFromProto` (method)

### dataTables

`add` (method)

### dataValidations

`items` (getter), `getForAddress` (method), `add` (method), `clear` (method)

### conditionalFormats

`addColorScale` (method), `addDataBar` (method), `addIconSet` (method), `addExpression` (method), `addCustom` (method), `addCellIs` (method), `add` (method), `items` (getter), `getItemAt` (method), `clearAll` (method), `deleteAll` (method), `clear` (method)

### freezePanes

`state` (getter), `freezeRows` (method), `freezeColumns` (method), `unfreeze` (method)


## Complete saved-source inventory

The snapshots are the dated reading basis. SHA-256 hashes are in the manifest; abbreviated hashes below help identify versions. Supporting experiment scripts and raw outputs are in the adjacent folders.

| Source | Bytes | SHA-256 prefix |
| --- | --- | --- |
| [package/API_QUICK_START.md](files/sources/package/API_QUICK_START.md) | 27041 | `d012a51f999b48ea` |
| [package/api/API_DOCS.md](files/sources/package/api/API_DOCS.md) | 8011 | `d0278d01fc56085a` |
| [package/api/references/charts-drawings.spec.md](files/sources/package/api/references/charts-drawings.spec.md) | 1728 | `26f913c6db2ff204` |
| [package/api/references/comments-notes-names.spec.md](files/sources/package/api/references/comments-notes-names.spec.md) | 1499 | `d581d13c4216ea44` |
| [package/api/references/formatting.spec.md](files/sources/package/api/references/formatting.spec.md) | 1762 | `0aaf7ca599d1ab37` |
| [package/api/references/formulas.spec.md](files/sources/package/api/references/formulas.spec.md) | 1511 | `89c75091b1da1485` |
| [package/api/references/import-export.md](files/sources/package/api/references/import-export.md) | 3387 | `a5a7ef8af9ec95d3` |
| [package/api/references/inspect-help.md](files/sources/package/api/references/inspect-help.md) | 2491 | `90d2d4399186a57b` |
| [package/api/references/ranges.spec.md](files/sources/package/api/references/ranges.spec.md) | 2915 | `12e84b016baf2d06` |
| [package/api/references/tables.spec.md](files/sources/package/api/references/tables.spec.md) | 1367 | `8e7503477413ddb6` |
| [package/api/references/workbook.spec.md](files/sources/package/api/references/workbook.spec.md) | 2641 | `203782076c26ff1a` |
| [package/api/references/worksheets.spec.md](files/sources/package/api/references/worksheets.spec.md) | 2295 | `5295185292d99539` |
| [package/examples/chart_suggestions.ts](files/sources/package/examples/chart_suggestions.ts) | 3567 | `ffba12b53f41de1a` |
| [package/examples/formula_trace_and_help.ts](files/sources/package/examples/formula_trace_and_help.ts) | 5813 | `c1d6767c52f5ce45` |
| [package/examples/inspect_existing_workbooks.ts](files/sources/package/examples/inspect_existing_workbooks.ts) | 3174 | `1daf6dab37693019` |
| [package/examples/quick_start_example.ts](files/sources/package/examples/quick_start_example.ts) | 4889 | `a0e23ee3e8d26215` |
| [package/formulas/database.md](files/sources/package/formulas/database.md) | 997 | `43c8f57654b669cf` |
| [package/formulas/date-time.md](files/sources/package/formulas/date-time.md) | 867 | `afbcd76848edebb4` |
| [package/formulas/engineering.md](files/sources/package/formulas/engineering.md) | 929 | `5694b05715bfabf7` |
| [package/formulas/financial.md](files/sources/package/formulas/financial.md) | 980 | `94ec433f2d11023b` |
| [package/formulas/information.md](files/sources/package/formulas/information.md) | 792 | `378ae765bff52f81` |
| [package/formulas/logical.md](files/sources/package/formulas/logical.md) | 1327 | `9d1a63c18b4380ad` |
| [package/formulas/lookup-reference.md](files/sources/package/formulas/lookup-reference.md) | 1022 | `da77855fcf5821dc` |
| [package/formulas/math-trig.md](files/sources/package/formulas/math-trig.md) | 868 | `bb13bd03018234f3` |
| [package/formulas/statistical.md](files/sources/package/formulas/statistical.md) | 881 | `ed2bf93edad916d8` |
| [package/formulas/text.md](files/sources/package/formulas/text.md) | 905 | `f3a2295e5cdf2691` |
| [package/references/comments.md](files/sources/package/references/comments.md) | 1022 | `3232ba2b04aa3ed6` |
| [package/references/conditional-formatting.spec.md](files/sources/package/references/conditional-formatting.spec.md) | 2433 | `79b16558a2d46d08` |
| [package/references/data-tables.spec.md](files/sources/package/references/data-tables.spec.md) | 750 | `d5d38dc1cab4d633` |
| [package/references/data-validations.spec.md](files/sources/package/references/data-validations.spec.md) | 1425 | `145ccd483e9133fa` |
| [package/references/defined-names.spec.md](files/sources/package/references/defined-names.spec.md) | 1031 | `088ae2855cb97019` |
| [package/references/drawings.spec.md](files/sources/package/references/drawings.spec.md) | 3542 | `2f16b38b57f6318e` |
| [package/references/enums.md](files/sources/package/references/enums.md) | 9180 | `25e0bf5b30fd829c` |
| [package/references/images.spec.md](files/sources/package/references/images.spec.md) | 1418 | `0a4a0c5696f9c821` |
| [package/references/ops/chart.md](files/sources/package/references/ops/chart.md) | 1666 | `bcc162d1bf010939` |
| [package/references/ops/comments.md](files/sources/package/references/ops/comments.md) | 419 | `abc0439a2550b3b8` |
| [package/references/ops/conditionalformat.md](files/sources/package/references/ops/conditionalformat.md) | 1039 | `0998627dafabe944` |
| [package/references/ops/datavalidation.md](files/sources/package/references/ops/datavalidation.md) | 1026 | `a04958ddfc05526e` |
| [package/references/ops/image.md](files/sources/package/references/ops/image.md) | 679 | `714b22ed45aceb56` |
| [package/references/ops/names.md](files/sources/package/references/ops/names.md) | 1106 | `ec1a7f26dad28c84` |
| [package/references/ops/range.md](files/sources/package/references/ops/range.md) | 2435 | `fed930c37708ad18` |
| [package/references/ops/shape.md](files/sources/package/references/ops/shape.md) | 986 | `10ee57ba435ff25b` |
| [package/references/ops/sheet.md](files/sources/package/references/ops/sheet.md) | 317 | `fbd9808a9bfb1032` |
| [package/references/ops/sparkline.md](files/sources/package/references/ops/sparkline.md) | 1358 | `3806a26f3a0000d6` |
| [package/references/ops/table.md](files/sources/package/references/ops/table.md) | 951 | `6e9a8d990c4df486` |
| [package/references/ops/thread.md](files/sources/package/references/ops/thread.md) | 1253 | `f62a03b3c3e23e94` |
| [package/references/ranges.spec.md](files/sources/package/references/ranges.spec.md) | 1710 | `e109cfa010b9a610` |
| [package/references/sparklines.spec.md](files/sources/package/references/sparklines.spec.md) | 1592 | `3aa83e6935c8c203` |
| [package/references/styles.spec.md](files/sources/package/references/styles.spec.md) | 2852 | `f04049b540df693c` |
| [package/references/tables.spec.md](files/sources/package/references/tables.spec.md) | 843 | `552fc166d6bdb2f3` |
| [package/references/workbook.spec.md](files/sources/package/references/workbook.spec.md) | 939 | `018df5de3a6bed03` |
| [skill/SKILL.md](files/sources/skill/SKILL.md) | 64729 | `499172cadf77be41` |
| [skill/agents/openai.yaml](files/sources/skill/agents/openai.yaml) | 364 | `8dc6dfa3a1c6fd33` |
| [skill/artifact_tool_docs/API_QUICK_START.md](files/sources/skill/artifact_tool_docs/API_QUICK_START.md) | 25830 | `2795943c6a356a6b` |
| [skill/artifact_tool_docs/DATA_TABLES.md](files/sources/skill/artifact_tool_docs/DATA_TABLES.md) | 3586 | `f87542406b8a2282` |
| [skill/artifact_tool_docs/SPARKLINES.md](files/sources/skill/artifact_tool_docs/SPARKLINES.md) | 1646 | `d3926fac25653829` |
| [skill/container_tools/mark_artifact_operation_started.mjs](files/sources/skill/container_tools/mark_artifact_operation_started.mjs) | 743 | `20c5a86e596154c0` |
| [skill/domain_guidance/financial_models.md](files/sources/skill/domain_guidance/financial_models.md) | 6987 | `ff163c47d4bcb14a` |
| [skill/domain_guidance/healthcare.md](files/sources/skill/domain_guidance/healthcare.md) | 4695 | `df5b9dcb9df988c1` |
| [skill/domain_guidance/marketing_advertising.md](files/sources/skill/domain_guidance/marketing_advertising.md) | 3913 | `bd6a715ac7d9ad2c` |
| [skill/domain_guidance/scientific_research.md](files/sources/skill/domain_guidance/scientific_research.md) | 3695 | `a95381e624c4c956` |
| [skill/features/charts.md](files/sources/skill/features/charts.md) | 5822 | `f85e0fa9ba2505f9` |
| [skill/references/image-references.md](files/sources/skill/references/image-references.md) | 1068 | `148690f09a1f33a8` |
| [skill/references/read_only_qna.md](files/sources/skill/references/read_only_qna.md) | 1095 | `687beb1ce4c32164` |
| [skill/references/template-elicitation.md](files/sources/skill/references/template-elicitation.md) | 1492 | `60b0ea3e373eef4f` |
| [skill/routing/google_sheets.md](files/sources/skill/routing/google_sheets.md) | 1456 | `0acafc1a346240b9` |
| [skill/style_guidelines.md](files/sources/skill/style_guidelines.md) | 10103 | `52f91ee6bc4c0a3b` |
| [skill/workflows/create_workflows.md](files/sources/skill/workflows/create_workflows.md) | 1563 | `92e3d0910bde06d2` |
| [skill/workflows/edit_workflows.md](files/sources/skill/workflows/edit_workflows.md) | 2481 | `f75917488c75b5a7` |
| [runtime/artifact-session/service.mjs](files/sources/runtime/artifact-session/service.mjs) | 30352 | `1eddcc994a6c5451` |
| [runtime/artifact-session/host-client.mjs](files/sources/runtime/artifact-session/host-client.mjs) | 18296 | `bece46d402a5a42b` |
| [runtime/artifact-session-mcp/server.mjs](files/sources/runtime/artifact-session-mcp/server.mjs) | 12783 | `af33172225f86354` |
| [runtime/artifact-session-mcp/worker.mjs](files/sources/runtime/artifact-session-mcp/worker.mjs) | 6190 | `7f4ae870c77a88de` |

## Final scope boundary

The finite documented spreadsheet catalogue and supporting workflows have been mapped and exercised as described. Service-dependent adapters, native Excel application behavior, every option combination and undocumented internal machinery are not certified. The deliverable is this cohesive Markdown reference plus reproducible evidence, rather than a claim of universal Excel equivalence.


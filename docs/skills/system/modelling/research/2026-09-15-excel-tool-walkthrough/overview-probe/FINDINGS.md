# What the overview actually returns

Read-only follow-up on 15 September 2026, using the same source.xlsx and installed Artifact Tool runtime. The original workbook and the page’s earlier saved captures were not changed.

The example labelled “Workbook overview” requests workbook, sheet and table records. With enough response allowance, its table records contain every value in each sheet’s reported used rectangle. All six returned matrices exactly matched direct range.values reads. This is structural data extraction, not a semantic summary of the model.

## Preview limits are conditional

The public workbook.inspect help says tableMaxRows/tableMaxCols/tableMaxCellChars control preview sizes “when values would be too large”. The experiments support that condition:

| Request | Returned data |
| --- | --- |
| Full workbook, 150,000-character allowance, 4 × 6 preview requested | All six full value matrices; 70,074 characters / approximately 22.4k o200k_base tokens |
| Fee engine table, 150,000-character allowance, 1 × 1 preview requested | Full 109 × 23 values; 24,242 characters / 9,533 tokens |
| Same Fee engine request, 1,000-character allowance | 1 × 1 preview; valuesTruncated: true, valuesPreviewAddress: A1 |
| Same small allowance with explicit range and include: values | Same 1 × 1 preview |
| Workbook and sheet records only | All six sheet descriptions; 642 characters / approximately 0.25k tokens |
| Fee engine regions, 2 × 3 previews | Six regions, every preview exactly 2 × 3, text capped at 20 characters; 1,707 characters / 559 tokens |

Token counts measure saved NDJSON with o200k_base, not exact GPT-6 billing.

Correction to the earlier interpretation: the table preview settings are not wholly ignored. They are fallback limits, not unconditional caps. This distinction matters when judging the API’s behavior.

## Why the original response stopped after the Assumptions description

Summary’s full table record is 6,602 characters; Assumptions’ is 10,062. Each is below the original 14,000-character allowance, but both together exceed it. The captured result contains the first table and the next sheet description, followed by a notice that nine lines were omitted. This is consistent with deciding whether to preview each table separately, then applying a response-wide limit to complete records. That mechanism is inferred from public help and observed outputs, not a reading of package internals.

The earlier response’s broad classification and truncation statements remain correct. Calling that output a compact overview was misleading. A useful compact first pass is workbook/sheet descriptions, followed by region previews or explicitly selected ranges when data is needed.

Evidence: [inputs and measurements](report.json), [public help](inspect-help.txt), [large table request](fee-table-tiny-preview.json), [small-budget request](fee-table-small-budget.json), [compact sheet overview](metadata-overview.json). The read-only script is [overview-probe.mjs](../overview-probe.mjs).

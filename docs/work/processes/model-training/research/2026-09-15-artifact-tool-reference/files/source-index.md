# Source inventory

## files/sources/package/API_QUICK_START.md

# Using artifact_tool APIs (JavaScript)
## Required Imports + Startup
## Build Patterns
## Conventions
## API Discovery Policy (Strict)
## Supported Formulas
### Known formula/export limitations (not exhaustive)
## Reading existing/imported workbooks
### Inspect for workbook understanding
## Known Gotchas (Do not repeat)
## Additional feature-specific notes
### Merging cells
## Quick API Surface (High-Value + Common)
### Core workbook/file APIs
### Worksheet selection/creation
### Worksheet operations
### Range values/formulas
### Formatting
### Data Validation
### Conditional formatting
### Tables
### Images
### Threaded Comments
### Charts
#### Fast-chart path
### Sparklines
### Help / Grep
### Trace
### JavaScript example snippet (runnable)

## files/sources/package/api/API_DOCS.md

# Workbook API Docs
## Conventions
## Quick Start
## Load Existing Workbook Data
## Recorded Edit
## Core API Sequence
## Output Map
## Minimal Patterns
## Reference Map

## files/sources/package/api/references/charts-drawings.spec.md

# Charts And Drawings
## Charts
## Chart To Image
## Images
## Shapes
## Sparklines
## Auto Layout
## Cookbook

## files/sources/package/api/references/comments-notes-names.spec.md

# Comments, Notes, And Names
## Comments
## Notes
## Workbook-Scoped Names
## Functions
## Sheet-Scoped Names
## Alias
## Cookbook

## files/sources/package/api/references/formatting.spec.md

# Formatting, Validation, And Theme
## Range Format
## Number Format
## Conditional Formatting
## Data Validation
## Freeze Panes
## Theme
## Cookbook

## files/sources/package/api/references/formulas.spec.md

# Formulas
## Assign Formulas
## Dynamic Arrays And Spill
## Trace
## Formula Usage Stats
## Names
## Cookbook

## files/sources/package/api/references/import-export.md

# Import, Export, HTML, Images, And Google Sheets
## CSV
## Markdown Table
## HTML
## Image Import
### Worksheet Images
## Chart To Image
## Render And Export
## Export Inline Types
## Google Sheets
## Cookbook

## files/sources/package/api/references/inspect-help.md

# Inspect, Resolve, Find Cells, And Help
## Inspect
## Targeted Inspect
## Resolve
## Find Cells
## Help
## Create/Edit Loop
## Inline Types

## files/sources/package/api/references/ranges.spec.md

# Ranges
## Addressing
## Values
## Formulas
## Write
## Format
## Navigation
## Clear, Merge, Fill, Copy
## Inline Types
## Cookbook

## files/sources/package/api/references/tables.spec.md

# Tables
## Add Table
## Read
## Add Rows
## Headerless Table
## Recorded Table Edit
## Cookbook

## files/sources/package/api/references/workbook.spec.md

# Workbook Facade
## Create And Load
## Import
## Root Collections
## Theme
## Recalculate And Trace
## Inspect, Search, Help, Resolve
## Record And Apply
## CRDT
## Serialize
## Utilities
## Cookbook

## files/sources/package/api/references/worksheets.spec.md

# Worksheets
## Collection
## Worksheet Properties
## Ranges
## Worksheet Collections
## Reset And Delete
## Merge
## Drawings
## Inline Types
## Cookbook

## files/sources/package/examples/chart_suggestions.ts



## files/sources/package/examples/formula_trace_and_help.ts



## files/sources/package/examples/inspect_existing_workbooks.ts



## files/sources/package/examples/quick_start_example.ts



## files/sources/package/formulas/database.md

# Database Formulas
## Formula Pattern
## Common Shapes
## Coverage

## files/sources/package/formulas/date-time.md

# Date And Time Formulas
## Formula Pattern
## Common Shapes
## Coverage

## files/sources/package/formulas/engineering.md

# Engineering Formulas
## Formula Pattern
## Common Shapes
## Coverage

## files/sources/package/formulas/financial.md

# Financial Formulas
## Formula Pattern
## Common Shapes
## Coverage

## files/sources/package/formulas/information.md

# Information Formulas
## Formula Pattern
## Common Shapes
## Coverage

## files/sources/package/formulas/logical.md

# Logical Formulas
## Formula Pattern
## Boolean And Branching Shapes
## Error Handling Shapes
## Lambda Shapes
## Named Lambda
## Coverage

## files/sources/package/formulas/lookup-reference.md

# Lookup And Reference Formulas
## Formula Pattern
## Common Shapes
## Coverage

## files/sources/package/formulas/math-trig.md

# Math And Trig Formulas
## Formula Pattern
## Common Shapes
## Coverage

## files/sources/package/formulas/statistical.md

# Statistical Formulas
## Formula Pattern
## Common Shapes
## Coverage

## files/sources/package/formulas/text.md

# Text Formulas
## Formula Pattern
## Common Shapes
## Coverage

## files/sources/package/references/comments.md

# Comments
## Current Author
## Person Inline Type
## Cell And Range Threads
## Thread Inline Types
## Replies And State

## files/sources/package/references/conditional-formatting.spec.md

# Conditional Formatting
## Rule Pattern
## Rule Inline Types
## Custom Rule
## Rule Families
## Clear Rules

## files/sources/package/references/data-tables.spec.md

## Data tables (`worksheet.dataTables`)
### Two-variable data table
### Single-variable (row input only)

## files/sources/package/references/data-validations.spec.md

# Data Validations
## Range Assignment
## Rule Config
## Validation Inline Types
## Collection Add

## files/sources/package/references/defined-names.spec.md

## Defined names (`workbook.names`, `worksheet.names`)
### Named ranges
### Named functions (LAMBDA)

## files/sources/package/references/drawings.spec.md

# Worksheet Drawings
## Units
## Chart Config
## Chart Inline Types
## Chart From Range
## Shape Config
## Shape Inline Type
## Rebuild And Layout
## Auto Layout Inline Type

## files/sources/package/references/enums.md

# Workbook.apply enums (v0)
## `props.axis.maxMode`
## `props.axis.minMode`
## `props.borders.preset`
## `props.chartType`
## `props.displayBlanksAs`
## `props.displayEmptyCellsAs`
## `props.errorAlert.style`
## `props.fill.color.value`
## `props.fill.gradientKind`
## `props.fill.pattern.type`
## `props.geometry`
## `props.horizontalAlignment`
## `props.legend.position`
## `props.line.style`
## `props.rule.format.fill.color.value`
## `props.rule.format.fill.gradientKind`
## `props.rule.format.fill.pattern.type`
## `props.rule.operator`
## `props.rule.thresholds[].type`
## `props.rule.type`
## `props.series[].marker.symbol`
## `props.series[].stroke.fill.color.value`
## `props.series[].stroke.fill.gradientKind`
## `props.series[].stroke.fill.pattern.type`
## `props.series[].stroke.style`
## `props.seriesColor.value`
## `props.type`
## `props.verticalAlignment`

## files/sources/package/references/images.spec.md

# Worksheet Images
## Add Image
## Image Inline Types
## Anchor
## Anchor Inline Type
## Edit

## files/sources/package/references/ops/chart.md

# chart ops
## Ops index
## `chart.add`

## files/sources/package/references/ops/comments.md

# comments ops
## Ops index
## `comments.self.set`

## files/sources/package/references/ops/conditionalformat.md

# conditionalformat ops
## Ops index
## `conditionalformat.add`

## files/sources/package/references/ops/datavalidation.md

# datavalidation ops
## Ops index
## `datavalidation.set`

## files/sources/package/references/ops/image.md

# image ops
## Ops index
## `image.add`

## files/sources/package/references/ops/names.md

# names ops
## Ops index
## `names.function.add`
## `names.range.add`

## files/sources/package/references/ops/range.md

# range ops
## Ops index
## `range.format.set`
## `range.formulas.set`
## `range.merge`
## `range.unmerge`
## `range.values.set`

## files/sources/package/references/ops/shape.md

# shape ops
## Ops index
## `shape.add`

## files/sources/package/references/ops/sheet.md

# sheet ops
## Ops index
## `sheet.add`

## files/sources/package/references/ops/sparkline.md

# sparkline ops
## Ops index
## `sparkline.add`

## files/sources/package/references/ops/table.md

# table ops
## Ops index
## `table.add`
## `table.rows.add`

## files/sources/package/references/ops/thread.md

# thread ops
## Ops index
## `thread.add`
## `thread.reopen`
## `thread.reply`
## `thread.resolve`

## files/sources/package/references/ranges.spec.md

## Ranges API
### Values
### Checkboxes
### Merged cells
### Formulas
### A1 addressing conventions

## files/sources/package/references/sparklines.spec.md

# Sparklines
## Add Group
## Sparkline Inline Type
## Range Alias
## Edit And Delete

## files/sources/package/references/styles.spec.md

# Range Formatting
## Shared Color And Fill Inline Types
## Grouped Format
## Format Inline Types
## Sizing
## Borders
## Border Inline Type

## files/sources/package/references/tables.spec.md

# Worksheet Tables
## Add Table
## Add Inline Type
## Rows And Values
## Table Style
## Table Inline Type
## Delete

## files/sources/package/references/workbook.spec.md

## Workbook API
### Quick start
### Workbook lifecycle
### Worksheets

## files/sources/skill/SKILL.md

# Spreadsheets skill
## Decision Boundary
## Important Instructions
## Tools + Contract Requirements
## Spreadsheet (Workbook) Complexity: Workbook Structure & Formulas
## Workbook Structure
### Tab Types & Relationships
### Tab Names
### Tab Order & Progression
#### Checks and Audit
### Build Structure and Formula Flow
### Workbook Structure Examples
## Formulas
### Formula Construction
### Anchoring
### Dates and Time Periods
### Choosing Formulas and Excel Tools
### Scalable Formulas and Brief Explanations
### Missing Inputs, Errors and Overrides
### Circular References and Iterative Calculation
### Formula Examples
## Writing Quality and Authored Content
## Workflows
## Resources
## Role and Domain Guidance
## Create and Edits
### Data Formatting Rules
### Verification Rules
### Citation Requirements
## Completion Criteria
### Criteria for Question / Read only requests
### Criteria for all create and edit requests
## Error Recovery
## Final response
### Final response citations
### Final response suggested followups
## Comment Author
## Source, PDF, and Attachment Processing

## files/sources/skill/agents/openai.yaml



## files/sources/skill/artifact_tool_docs/API_QUICK_START.md

# Using artifact_tool APIs (JavaScript)
## Required imports, setup and exports
## Build Patterns
## Conventions
## API Discovery Policy (Strict)
## Supported Formulas
### Known formula/export limitations (not exhaustive)
## Reading existing/imported workbooks
### Inspect for workbook understanding
## Known Gotchas (Do not repeat)
## Quick API Surface (High-Value + Common)
### Core workbook/file APIs
### Worksheet selection/creation
### Worksheet operations
### Range values/formulas
### Formatting
### Data Validation
### Conditional formatting
### Tables
### Images
### Threaded Comments
### Notes
### Charts
#### Fast-chart path
### Merging cells
### Help / Grep
### Trace
## Additional Resources
## JavaScript example snippet (runnable)

## files/sources/skill/artifact_tool_docs/DATA_TABLES.md

# Data tables (two-variable sensitivity)

## files/sources/skill/artifact_tool_docs/SPARKLINES.md

# Sparklines

## files/sources/skill/container_tools/mark_artifact_operation_started.mjs



## files/sources/skill/domain_guidance/financial_models.md

# Finance Guidance
## Financial Basis and Sources
## Periods, Assumptions and Scenarios
## Operating Schedules and Reporting
## Valuation and Returns
## Finance Presentation
## Finance Audit and Verification

## files/sources/skill/domain_guidance/healthcare.md

## Healthcare (clinical/administrative) spreadsheets
### Tab structure
### Formatting
### Formulas
### Raw data and outputs
### Metadata, units, and codes

## files/sources/skill/domain_guidance/marketing_advertising.md

## Marketing/Advertising Guidance
### Tab structure
### Cell formatting
### Marketing Analysis
### Raw data vs. outputs
### Metadata and sources

## files/sources/skill/domain_guidance/scientific_research.md

## Scientific Research Guidance
### Tab structure & naming
### Formatting
### Formula practices

## files/sources/skill/features/charts.md

# Charts
## Choose the chart
## Keep data auditable
## Format and place
## Verify and repair
## Edit existing charts

## files/sources/skill/references/image-references.md

## Importing or extracting data from screenshots or reference images

## files/sources/skill/references/read_only_qna.md

## Handling queries and questions

## files/sources/skill/references/template-elicitation.md

# Artifact Template Selection

## files/sources/skill/routing/google_sheets.md

## Google Sheets-targeted output
### New Creations
### Edits

## files/sources/skill/style_guidelines.md

# Style and Formatting Instructions
## Tab Structure Defaults
## Reader-facing sheet layout
## Use a visually clear layout
## Freeze panes
## Align and format by data type
## Typography - Use intentionally but conservatively
## Live inputs and visuals

## files/sources/skill/workflows/create_workflows.md

# Workflows for creating new spreadsheets
## Quality Guidelines
## Checks

## files/sources/skill/workflows/edit_workflows.md

# Editing an existing spreadsheet
## Safety rules
## Important guidelines
## Formula Rules

## files/sources/runtime/artifact-session/service.mjs



## files/sources/runtime/artifact-session/host-client.mjs



## files/sources/runtime/artifact-session-mcp/server.mjs



## files/sources/runtime/artifact-session-mcp/worker.mjs



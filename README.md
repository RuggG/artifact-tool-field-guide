# Artifact Tool: OpenAI's Modeling Engine

An independent, interactive guide to Artifact Tool for Excel: action families, recorded inputs and outputs, formula reference, spreadsheet skills and working practices.

**[Open the public guide](https://ruggg.github.io/artifact-tool-field-guide/)**

The complete guide, companion walkthrough, captured examples, supporting documentation and downloadable workbooks are served together. No Lynott Hub account is required. The controls explore saved executions; they do not run the workbook engine in the browser.

## Publication

GitHub Pages serves `docs/` from the `main` branch. `.nojekyll` preserves source files and paths as-is. The public edition keeps the original guide content and presentation while replacing Hub-specific navigation and services with standalone navigation. Markdown and source files have readable HTML views and raw/download links.

## Refresh the saved edition

Install `requirements.txt`, then run:

```sh
python tools/export.py --hub-root /path/to/lynott-hub --extra-root /path/to/other/source-checkout
```

The exporter reads source files without modifying the Hub. Run the public-site validation before committing and publishing refreshed files.

## Provenance

Prepared by Ruggero Gargiulo with Codex assistance. The guide records Artifact Tool version 2.8.59 and the source material captured in September 2026. Original documentation and example provenance remain with their files. This is an independent field guide, not official OpenAI documentation. Third-party source material retains its original attribution.

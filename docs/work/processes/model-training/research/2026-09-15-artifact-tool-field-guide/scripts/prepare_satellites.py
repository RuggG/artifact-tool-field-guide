"""Freeze the complete skill-document catalogue and reproducible token counts.

Run with a Python environment containing tiktoken, then render_satellites.mjs.
Counts cover each original UTF-8 Markdown file, including metadata and code.
"""
import hashlib
import importlib.metadata
import json
from pathlib import Path
import tiktoken

HOME = Path(__file__).resolve().parents[1]
SOURCE = HOME.parent / '2026-09-15-artifact-tool-reference/files/sources/skill'
GROUPS = [
    ('foundation', 'Structure, style and finance', 'Start here to understand how the agent builds and presents a workbook.'),
    ('task', 'Guidance for a particular task', 'Read when charts, references, templates or workbook questions are involved.'),
    ('technical', 'API and delivery references', 'Call patterns, specialist features and the Google Sheets handoff.'),
    ('other', 'Other domain guides', 'Included for reference; use only when their requirements fit the work.'),
]
ROWS = [
    ('main', 'foundation', 'Main spreadsheet skill', 'SKILL.md', 'The entry point: workbook structure, formulas, sources, verification and which supporting documents to read.', 'Start here; it routes the agent to the other documents.'),
    ('style', 'foundation', 'Style and layout', 'style_guidelines.md', 'Tab colours, sheet layout, typography, number formats, spacing, panes and editable-input cues.', 'Required formatting guidance; preserve user and reference conventions.'),
    ('finance', 'foundation', 'Finance guidance', 'domain_guidance/financial_models.md', 'Financial periods, assumptions, scenarios, operating schedules, presentation and audit checks.', 'For financial work in any industry; read the relevant requirements.'),
    ('create', 'foundation', 'Creating a workbook', 'workflows/create_workflows.md', 'Choose the audience, outputs and logical structure before adding sheets and content.', 'Required workflow for new files.'),
    ('edit', 'foundation', 'Editing a workbook', 'workflows/edit_workflows.md', 'Inspect the existing design, make focused changes and preserve unrelated content and native features.', 'Required workflow for existing files and follow-ups.'),
    ('charts', 'task', 'Charts', 'features/charts.md', 'Choose a useful chart, bind its data, set honest axes and labels, place it and verify export.', 'When creating or editing charts.'),
    ('images', 'task', 'Working from a screenshot', 'references/image-references.md', 'Recreate the intended appearance while retaining real numbers, dates and justified formulas.', 'When an image or screenshot is supplied as a reference.'),
    ('questions', 'task', 'Questions and audits', 'references/read_only_qna.md', 'Locate the right cells, follow formulas and reconcile an answer without changing the workbook.', 'For read-only questions and audits.'),
    ('templates', 'task', 'Choosing a template', 'references/template-elicitation.md', 'When to offer the template picker, how to use its choice and when to continue without a template.', 'When no template, reference or visual direction has been supplied, subject to the documented exceptions.'),
    ('api', 'technical', 'API Quick Start', 'artifact_tool_docs/API_QUICK_START.md', 'Imports, workbook calls, recommended patterns, examples and known limitations.', 'The main skill requires reading this API guide in full before starting the task.'),
    ('sensitivity', 'technical', 'What-If Data Tables', 'artifact_tool_docs/DATA_TABLES.md', 'Set up two-input sensitivity analysis; these are different from ordinary named Excel Tables.', 'Before creating a native What-If Data Table.'),
    ('sparklines', 'technical', 'Sparklines', 'artifact_tool_docs/SPARKLINES.md', 'Create and configure miniature charts inside cells.', 'Feature reference linked from the API Quick Start.'),
    ('google', 'technical', 'Google Sheets routing', 'routing/google_sheets.md', 'Choose the creation/import or editing workflow when the destination is Google Sheets.', 'When working with Google Sheets.'),
    ('healthcare', 'other', 'Healthcare', 'domain_guidance/healthcare.md', 'Clinical and administrative records, units, codes and supplied protocols.', 'Only when relevant to the task; a company\u2019s industry does not dictate the workbook design.'),
    ('marketing', 'other', 'Marketing and advertising', 'domain_guidance/marketing_advertising.md', 'Campaign definitions, metrics, original extracts and supported targets.', 'Only when relevant to the task.'),
    ('science', 'other', 'Scientific research', 'domain_guidance/scientific_research.md', 'Observations, units, reproducible processing and requested uncertainty or checks.', 'Only when relevant to the task.'),
]

def main():
    encoding = tiktoken.get_encoding('o200k_base')
    docs = []
    for id, group, title, path, summary, when in ROWS:
        source_bytes = (SOURCE / path).read_bytes()
        raw = source_bytes.decode('utf-8')
        docs.append(dict(id=id, group=group, title=title, path=path, summary=summary, when=when,
                         source='../2026-09-15-artifact-tool-reference/files/sources/skill/' + path,
                         markdown=raw, tokens=len(encoding.encode(raw, disallowed_special=())),
                         words=len(raw.split()), characters=len(raw), bytes=len(source_bytes),
                         sha256=hashlib.sha256(source_bytes).hexdigest()))
    expected = {str(p.relative_to(SOURCE)) for p in SOURCE.rglob('*.md')}
    assert {d['path'] for d in docs} == expected, 'Catalogue must contain every saved Markdown document.'
    result = dict(captured='2026-09-15', counted='2026-09-16', encoding=encoding.name,
                  tokenizer='tiktoken', tokenizerVersion=importlib.metadata.version('tiktoken'),
                  countScope='Complete raw Markdown, including front matter, code, links and whitespace. Counts are exact for o200k_base, not a model-specific prompt or billing measurement.',
                  groups=[dict(id=id, title=title, summary=summary) for id, title, summary in GROUPS],
                  documents=docs, totals={k: sum(d[k] for d in docs) for k in ('tokens','words','characters','bytes')})
    (HOME / 'data/satellite-documents.json').write_text(json.dumps(result, ensure_ascii=False, indent=2) + '\n')
    print(f'{len(docs)} documents; {result["totals"]["tokens"]:,} o200k_base tokens')

if __name__ == '__main__':
    main()

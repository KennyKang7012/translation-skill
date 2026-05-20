---
name: zh-tw-document-localization
description: Translate arbitrary user-provided PDF, Markdown, DOCX, or PPTX files into Taiwan Traditional Chinese while preserving layout, images, styling, structural anchors, and deterministic .zh-tw output filenames. Use for high-fidelity document localization, template-agnostic translation workflows, and QA-reportable source-to-target mapping.
---

# zh-tw Document Localization

Localize documents into Taiwan Traditional Chinese with a generic, format-adapter workflow. Treat every input as an unknown template: inspect structure first, preserve anchors, and write translated text back only through stable source positions.

## Runtime

Use `uv` for Python environment management.

1. Enter the skill folder: `cd zh-tw-document-localization`
2. Create the virtual environment: `uv venv`
3. Install and sync dependencies: `uv sync`
4. Set credentials: `OPENAI_API_KEY`
5. Optionally set default model: `OPENAI_MODEL`
6. Run deterministic localization: `uv run python scripts/localize.py <input_file>`
7. Run agent workflow with report: `uv run python scripts/agent_pipeline.py <input_file> --emit-report`
8. Limit PDF test scope: `uv run python scripts/agent_pipeline.py <input_file> --max-pages 3 --emit-report`
9. Compute the output path only: `uv run python scripts/rename_output.py <input_file> --print-only`

## Core Contract

Every format adapter must use the same logical pipeline:

1. Extract translatable units from the source file.
2. Assign a stable structural `anchor_id` to each unit.
3. Preserve nearby context, style, geometry, and locked tokens.
4. Translate only the unit text.
5. Reinsert translated text at the original anchor.
6. Validate structure, output filename, and localization completeness.

Use this generic record shape for all formats:

```text
anchor_id: source-position key, not derived from text alone
source_text: original text
target_text: translated zh-TW text
container_path: page/slide/paragraph/shape/cell path
style_hint: font, size, color, alignment, or syntax metadata
geometry: page rectangle, shape bounds, table cell, or null
flags: code, url, path, product_name, glossary_lock, non_translatable
```

## Universal Rules

1. Never hardcode behavior for a specific file, title, product, page number, template, or brand.
2. Never map translations by source text alone; repeated text must be resolved by `anchor_id`.
3. Never overwrite the source file.
4. Always output as `basename.zh-tw.ext`.
5. Preserve images, object positions, visual hierarchy, tables, links, and non-text assets.
6. Preserve locked tokens such as URLs, file paths, commands, code, API names, model names, product names, and glossary terms.
7. Prefer failing with a report over silently dropping text.

## Format Adapters

### Markdown

Translate prose while preserving Markdown structure. Keep code fences, inline code, frontmatter keys, links, image paths, table alignment markers, list indentation, and reference identifiers intact.

### DOCX

Translate text in place using document structure. Preserve paragraphs, runs when possible, styles, tables, headers, footers, section breaks, page numbers, comments, embedded media, and document relationships.

### PPTX

Translate shape, text frame, table cell, and speaker-note text in place. Preserve slide masters, layouts, theme colors, animations, grouping, z-order, object IDs, charts, images, and shape geometry.

### PDF

Use the least destructive route available.

1. If editable source exists, translate the editable source and export to PDF.
2. If only PDF is available, use page/block/line/span anchors.
3. Repaint removed text using sampled local background color, not a fixed fill.
4. Fit translated text inside the original bounding box when possible.
5. Record fallback placement, overflow, or missing-font behavior in the report.

## Translation Rules

1. Use natural Taiwan Traditional Chinese.
2. Keep product names and technical identifiers stable unless a glossary explicitly says otherwise.
3. Keep numbering, citations, references, formulas, URLs, CLI flags, and file paths unchanged.
4. Maintain terminology consistency across the document.
5. Use concise wording when the target container has limited space.

## QA Requirements

Generate or inspect a report for non-trivial files.

1. Verify extracted anchor count, translated count, skipped count, failed count, and QA revision count.
2. Confirm no required anchor was skipped without explanation.
3. Confirm filename follows `basename.zh-tw.ext`.
4. Confirm output opens successfully.
5. For visual formats, inspect at least a representative page or slide before treating the run as complete.

## Resources

Use `scripts/localize.py` for direct localization.  
Use `scripts/agent_pipeline.py` for coordinator/translator/QA reporting.  
Use `scripts/rename_output.py` only for deterministic filename generation or bootstrap copying.  
Read `references/format-notes.md` when choosing the safest implementation path for a file format.  
Read `references/universal-design.md` only before changing adapter contracts, anchor strategy, reinjection rules, or source-to-target mapping logic.

---
name: zh-tw-document-localization
description: Translate user-provided files into Taiwan Traditional Chinese while preserving original layout, image positions, and visual style. Use when handling PDF, Markdown, DOCX, or PPTX localization requests that require high-fidelity output and deterministic zh-tw filename suffixing.
---

# zh-tw Document Localization

Follow this workflow to localize documents into Taiwan Traditional Chinese without changing structure or design.

## Environment (uv)

Use `uv` as the only Python environment manager for this skill.

1. Create environment: `uv venv`
2. Install dependencies in project scope: `uv pip install -r requirements.txt`
3. Run localization with environment isolation: `uv run python scripts/localize.py <input_file>`
4. Set `OPENAI_API_KEY` before running translation.
5. Use `uv run python scripts/rename_output.py <input_file> --print-only` when only path computation is needed.

## Universal Architecture

Use a format-agnostic pipeline with per-format adapters.

1. Extract: collect translatable units and preserve structural metadata.
2. Segment: split text into stable translation units without breaking structure.
3. Anchor: assign immutable anchor IDs for each unit (page/slide/paragraph/run/cell).
4. Translate: produce zh-tw text per anchor ID.
5. Reinject: write translations back to the exact original anchor location.
6. Verify: run structure, visual, and completeness checks before output.

Keep this contract for all formats:

- `anchor_id`: unique position key in source structure
- `source_text`: original text
- `target_text`: translated zh-tw text
- `context`: nearby metadata for disambiguation
- `flags`: non-translatable markers (code, URL, path, term lock)

## Output Rules

1. Keep page layout, spacing, image coordinates, table geometry, font hierarchy, and color palette unchanged.
2. Translate text content only.
3. Name output files as `original_basename.zh-tw.ext`.
4. Never overwrite source files.

## Format Workflow

### Markdown (.md)

1. Preserve Markdown syntax, heading levels, code fences, links, image paths, and frontmatter keys.
2. Translate prose only; keep code blocks and inline code unchanged unless the user asks otherwise.
3. Preserve table alignment markers, list indentation depth, and reference-style link IDs.
3. Save as `*.zh-tw.md`.

### DOCX (.docx)

1. Edit existing paragraphs/runs in-place instead of recreating the document.
2. Preserve styles, section breaks, headers/footers, page numbers, comments, tracked elements, and embedded objects.
3. Do not delete/reinsert images or tables.
4. Preserve run boundaries when possible to avoid style leakage.
4. Save as `*.zh-tw.docx`.

### PPTX (.pptx)

1. Translate shape text in-place slide by slide.
2. Keep slide masters, theme colors, animations, object IDs, and z-order unchanged.
3. Do not move or resize images/charts/shapes.
4. Keep speaker notes, grouped shapes text order, and table cell coordinates unchanged.
4. Save as `*.zh-tw.pptx`.

### PDF (.pdf)

1. First choose a fidelity-first route:
   - If editable source exists, translate source format and export back to PDF.
   - If source is unavailable, use OCR/text-layer replacement while keeping original page geometry.
2. Keep the same page size, margins, image positions, and color appearance.
3. Preserve reading order anchors by page + block + line index.
3. Save as `*.zh-tw.pdf`.

## Precision Mapping Rules

1. Never map by text content alone; map by structural anchor first.
2. If duplicate source strings exist, distinguish by anchor path (page/slide/paragraph index).
3. Prevent overflow regressions:
   - allow line reflow only inside the original text container
   - never expand container dimensions automatically
4. If target text does not fit, apply constrained fallback in order:
   - terminology shortening
   - punctuation compaction
   - controlled font-size reduction within allowed tolerance
5. Track all fallback decisions in a change log section of run output.

## Translation Conventions (Taiwan)

1. Use Taiwan Traditional Chinese terms and punctuation.
2. Preserve product names, API names, commands, and file paths.
3. Keep numbering, references, URLs, and citation anchors stable.
4. Keep glossary consistency across the full file.

## QA Checklist

1. Visual diff check: no image displacement, no theme drift, no structural collapse.
2. Semantic check: no missing paragraphs, list items, table cells, or speaker notes.
3. Filename check: confirm exact `basename.zh-tw.ext` pattern.
4. Encoding check: ensure readable Traditional Chinese output.
5. Anchor completeness check: translated anchor count must equal extracted anchor count.

## Helper Script

Use `scripts/localize.py` for actual translation output.  
Use `scripts/rename_output.py` only for output naming or file bootstrap behavior.

Read `references/universal-design.md` before implementing new format handlers or changing mapping logic.

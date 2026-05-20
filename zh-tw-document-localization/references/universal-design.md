# Universal Design for Multi-Format Localization

This reference defines a generic, reusable design for PDF/Markdown/DOCX/PPTX localization with high fidelity.

## Goals

1. Preserve visual fidelity and structure.
2. Ensure deterministic source-to-target mapping.
3. Support template variance without format-specific hardcoding.

## Canonical Data Model

Represent extracted text as normalized records:

```text
record {
  anchor_id: string
  format: pdf|md|docx|pptx
  container_path: string
  source_text: string
  target_text: string|null
  context_before: string|null
  context_after: string|null
  style_hint: string|null
  flags: string[]
}
```

## Anchor Strategy

Use stable position-based anchors:

- PDF: `p{page}.b{block}.l{line}.s{span}`
- Markdown: `sec{heading_index}.blk{block_index}.ln{line_index}`
- DOCX: `sec{s}.p{p}.r{r}.t{t}`
- PPTX: `sl{slide}.sh{shape}.tf{textframe}.p{p}.r{r}`

Rules:

1. Keep anchors immutable through the whole pipeline.
2. Rebuild missing anchors only from original source, never from translated output.

## Translation Unit Rules

1. Keep sentence continuity but avoid crossing container boundaries.
2. Do not merge units from different anchors.
3. Lock non-translatable tokens with placeholders:
   - `{{CODE_1}}`
   - `{{URL_1}}`
   - `{{PATH_1}}`
   - `{{TERM_1}}`

## Reinjection Rules

1. Reinjection must target original anchor IDs only.
2. If an anchor cannot be written back, fail the run (do not silently skip).
3. Generate a mismatch report:
   - missing anchors
   - duplicate anchors
   - overflow anchors

## Validation Gates

Pass all gates before shipping output:

1. Structural parity: same object count and ordering constraints.
2. Anchor parity: extracted count equals injected count.
3. Visual parity: no moved images, no palette drift, no shape resize.
4. Text quality: no source-language leakage except locked tokens.

## Extensibility

To add a new format, only implement:

1. Extractor adapter
2. Anchor generator
3. Reinjector adapter
4. Validator adapter

Keep the canonical data model and gates unchanged.

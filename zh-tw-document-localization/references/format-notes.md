# Format Notes

Use this quick guide when selecting an implementation path.

## Runtime Standard

- Use `uv` to manage Python runtime and dependencies.
- Prefer `uv run ...` instead of direct `python ...` execution.

## Priority

1. In-place editing path (best fidelity)
2. Source-format round trip (acceptable)
3. OCR/reflow rebuild (last resort)

## Important Guardrails

- Do not rebuild DOCX/PPTX from plain text if preserving visual fidelity is required.
- For PDF, prefer translating original source document then exporting to PDF.
- Keep non-translatable tokens unchanged:
  - code
  - CLI flags
  - URLs
  - file paths
  - proper nouns

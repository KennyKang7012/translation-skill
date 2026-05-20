# Taiwan Traditional Chinese Document Localization Skill

This repository contains a Codex skill package for localizing user-provided documents into Taiwan Traditional Chinese (`zh-TW`) while preserving as much of the original layout as possible.

Supported input formats:

- PDF
- Markdown
- DOCX
- PPTX

Localized output files use this deterministic filename pattern:

```text
<original-basename>.zh-tw.<extension>
```

Example:

```text
The-Complete-Guide-to-Building-Skill-for-Claude.pdf
The-Complete-Guide-to-Building-Skill-for-Claude.zh-tw.pdf
```

## Repository Layout

```text
zh-tw-document-localization/
  SKILL.md
  pyproject.toml
  requirements.txt
  uv.lock
  .env.example
  agents/
    openai.yaml
  scripts/
    agent_pipeline.py
    localize.py
    rename_output.py
  references/
    format-notes.md
    universal-design.md
```

## Setup

Install `uv` first.

From a fresh clone, enter the skill folder:

```powershell
cd D:\VibeCoding\translation-skill\zh-tw-document-localization
```

Create a Python virtual environment:

```powershell
uv venv
```

Install and sync dependencies:

```powershell
uv sync
```

## Environment Variables

Create a local `.env` file from the example:

```powershell
copy .env.example .env
```

Set at least:

```text
OPENAI_API_KEY=your_openai_api_key_here
```

Optional settings:

```text
OPENAI_MODEL=gpt-4.1-mini
PDF_CJK_FONT_PATH=C:\Windows\Fonts\msjh.ttc
```

If `PDF_CJK_FONT_PATH` is not set, the PDF writer tries these fonts in order:

- `C:\Windows\Fonts\msjh.ttc`
- `C:\Windows\Fonts\msjhbd.ttc`
- `C:\Windows\Fonts\NotoSansTC-VF.ttf`
- PyMuPDF built-in CJK font fallback

## Usage

The agent pipeline is the recommended entrypoint because it also emits a report.

```powershell
uv run python scripts/agent_pipeline.py <input_file> --emit-report
```

Example:

```powershell
uv run python scripts/agent_pipeline.py ..\docs\example.pdf --emit-report
```

For large PDFs, test only the first few pages first:

```powershell
uv run python scripts/agent_pipeline.py ..\docs\example.pdf --max-pages 3 --emit-report
```

To run faster and avoid the QA rewrite pass:

```powershell
uv run python scripts/agent_pipeline.py ..\docs\example.pdf --disable-qa --emit-report
```

Direct localization without the agent report workflow:

```powershell
uv run python scripts/localize.py <input_file>
```

Only compute the output filename:

```powershell
uv run python scripts/rename_output.py <input_file> --print-only
```

## PDF Strategy

PDF is the hardest format to preserve perfectly because the script must write translated text back into existing visual regions.

The current PDF strategy is conservative:

- Preserve images, backgrounds, code blocks, tables, and original page geometry.
- Use block-level context for heading-like text.
- Use line-level translation for body text to reduce overlap risk.
- Skip high-risk code/table/tree-like text when translating it would damage layout.
- Keep source text when translated text cannot safely fit into the original text box.
- Record skipped or failed segments in the report.
- Prefer system Traditional Chinese fonts such as Microsoft JhengHei.

This means the output prioritizes layout safety. Some text may remain untranslated when fitting it back would cause overlap or visual damage.

## Reports

When `--emit-report` is used, the pipeline writes a report next to the output file:

```text
<output-file>.report.txt
```

The report includes:

- page count
- extracted segment count
- translated segment count
- skipped segment count
- failed segment count
- QA pass and revision counts
- per-page status
- failure reasons

## Skill Files

The main skill definition is:

```text
zh-tw-document-localization/SKILL.md
```

Reference documents:

- `references/format-notes.md`: format strategy notes
- `references/universal-design.md`: adapter, anchor, and reinjection design

## Notes

- `.env`, `docs/`, `__pycache__/`, and `*.pyc` are ignored by git.
- If the output PDF is open in a PDF viewer, PyMuPDF may fail to overwrite it. Close the viewer before rerunning.
- If an editable source file exists, prefer translating that source and exporting back to PDF. PDF-only localization is a best-effort fallback.

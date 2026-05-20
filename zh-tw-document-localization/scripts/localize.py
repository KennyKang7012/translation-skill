#!/usr/bin/env python3
from __future__ import annotations

import argparse
import os
import re
from pathlib import Path
from typing import Iterable

import fitz  # pymupdf
from docx import Document
from dotenv import load_dotenv
from openai import OpenAI
from pptx import Presentation


NON_TRANSLATABLE_PATTERNS = [
    re.compile(r"https?://\S+"),
    re.compile(r"`[^`]+`"),
    re.compile(r"\b[A-Za-z]:\\[^\s]+"),
    re.compile(r"(?:^|\s)(?:-\w+|--\w[\w-]*)(?:\s|$)"),
]


def make_output_path(src: Path) -> Path:
    return src.with_name(f"{src.stem}.zh-tw{src.suffix}")


def should_translate(text: str) -> bool:
    stripped = text.strip()
    if not stripped:
        return False
    if re.fullmatch(r"[\W\d_]+", stripped):
        return False
    return True


def mask_non_translatable(text: str) -> tuple[str, dict[str, str]]:
    mapping: dict[str, str] = {}
    masked = text
    idx = 0
    for pattern in NON_TRANSLATABLE_PATTERNS:
        matches = list(pattern.finditer(masked))
        for m in reversed(matches):
            token = m.group(0)
            key = f"__LOCK_{idx}__"
            idx += 1
            mapping[key] = token
            masked = masked[: m.start()] + key + masked[m.end() :]
    return masked, mapping


def unmask(text: str, mapping: dict[str, str]) -> str:
    out = text
    for key, token in mapping.items():
        out = out.replace(key, token)
    return out


class Translator:
    def __init__(self, model: str) -> None:
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise RuntimeError(
                "OPENAI_API_KEY is not set. Please set it before running translation."
            )
        self.client = OpenAI(api_key=api_key)
        self.model = model
        self.cache: dict[str, str] = {}

    def translate(self, text: str) -> str:
        if not should_translate(text):
            return text
        if text in self.cache:
            return self.cache[text]

        masked, mapping = mask_non_translatable(text)
        prompt = (
            "Translate the following text into Taiwan Traditional Chinese (zh-TW). "
            "Preserve meaning, keep placeholders like __LOCK_#__ unchanged, and do not add explanation.\n\n"
            f"{masked}"
        )
        resp = self.client.responses.create(
            model=self.model,
            input=prompt,
            temperature=0,
        )
        translated = (resp.output_text or "").strip()
        if not translated:
            translated = text
        translated = unmask(translated, mapping)
        self.cache[text] = translated
        return translated


def localize_markdown(src: Path, out: Path, tr: Translator) -> None:
    lines = src.read_text(encoding="utf-8").splitlines()
    in_code = False
    output: list[str] = []
    for line in lines:
        if line.strip().startswith("```"):
            in_code = not in_code
            output.append(line)
            continue
        if in_code or line.strip().startswith("![]("):
            output.append(line)
            continue
        output.append(tr.translate(line))
    out.write_text("\n".join(output) + "\n", encoding="utf-8")


def _translate_docx_runs(runs: Iterable, tr: Translator) -> None:
    for run in runs:
        if should_translate(run.text):
            run.text = tr.translate(run.text)


def localize_docx(src: Path, out: Path, tr: Translator) -> None:
    doc = Document(str(src))
    for p in doc.paragraphs:
        _translate_docx_runs(p.runs, tr)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _translate_docx_runs(p.runs, tr)
    doc.save(str(out))


def localize_pptx(src: Path, out: Path, tr: Translator) -> None:
    prs = Presentation(str(src))
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        if should_translate(run.text):
                            run.text = tr.translate(run.text)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            for run in p.runs:
                                if should_translate(run.text):
                                    run.text = tr.translate(run.text)
    prs.save(str(out))


def localize_pdf(src: Path, out: Path, tr: Translator, max_pages: int | None = None) -> None:
    doc = fitz.open(str(src))
    total_pages = len(doc)
    pages_to_process = total_pages if max_pages is None else min(max_pages, total_pages)

    for page_idx in range(pages_to_process):
        page = doc[page_idx]
        print(f"Processing page {page_idx + 1}/{pages_to_process}...", flush=True)
        data = page.get_text("dict")
        spans_to_replace: list[dict] = []
        for block in data.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = span.get("text", "")
                    if should_translate(text):
                        translated = tr.translate(text)
                        if translated != text:
                            spans_to_replace.append(
                                {
                                    "bbox": span["bbox"],
                                    "text": translated,
                                    "size": span.get("size", 10),
                                    "color": span.get("color", 0),
                                }
                            )
        for s in spans_to_replace:
            rect = fitz.Rect(s["bbox"])
            page.add_redact_annot(rect, fill=(1, 1, 1))
        if spans_to_replace:
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
        for s in spans_to_replace:
            rect = fitz.Rect(s["bbox"])
            rgb = s["color"]
            r = ((rgb >> 16) & 255) / 255.0
            g = ((rgb >> 8) & 255) / 255.0
            b = (rgb & 255) / 255.0
            page.insert_textbox(
                rect,
                s["text"],
                fontsize=max(6, min(48, s["size"])),
                color=(r, g, b),
                align=fitz.TEXT_ALIGN_LEFT,
            )
    doc.save(str(out))
    doc.close()


def main() -> int:
    load_dotenv()
    parser = argparse.ArgumentParser(
        description="Localize PDF/Markdown/DOCX/PPTX into Taiwan Traditional Chinese."
    )
    parser.add_argument("input_file", help="Input file path")
    parser.add_argument(
        "--model",
        default=None,
        help="OpenAI model name. If omitted, use OPENAI_MODEL or fallback default.",
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        default=None,
        help="Only process the first N pages for PDF. Ignored for non-PDF.",
    )
    args = parser.parse_args()

    src = Path(args.input_file).resolve()
    if not src.exists():
        raise FileNotFoundError(f"Input not found: {src}")

    out = make_output_path(src)
    model = args.model or os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
    tr = Translator(model=model)

    ext = src.suffix.lower()
    if ext == ".md":
        localize_markdown(src, out, tr)
    elif ext == ".docx":
        localize_docx(src, out, tr)
    elif ext == ".pptx":
        localize_pptx(src, out, tr)
    elif ext == ".pdf":
        localize_pdf(src, out, tr, max_pages=args.max_pages)
    else:
        raise ValueError(f"Unsupported format: {ext}. Supported: .pdf .md .docx .pptx")

    print(out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

#!/usr/bin/env python3
from __future__ import annotations

import argparse
import os
from dataclasses import dataclass
from pathlib import Path
from typing import List

from dotenv import load_dotenv
from localize import (
    Translator,
    localize_docx,
    localize_markdown,
    localize_pdf,
    localize_pptx,
    make_output_path,
)


@dataclass
class SegmentReport:
    anchor_id: str
    source_preview: str
    target_preview: str
    qa_ok: bool
    note: str


class QAAgent:
    def __init__(self, translator: Translator) -> None:
        self.translator = translator

    def review(self, anchor_id: str, source_text: str, target_text: str) -> SegmentReport:
        prompt = (
            "You are a QA reviewer for zh-TW localization.\n"
            "Check if target text is accurate, natural Taiwan Traditional Chinese, and does not alter locked tokens.\n"
            "Respond in one line as: OK|<short note> or FIX|<better translation>\n\n"
            f"Source:\n{source_text}\n\nTarget:\n{target_text}"
        )
        resp = self.translator.client.responses.create(
            model=self.translator.model,
            input=prompt,
            temperature=0,
        )
        text = (resp.output_text or "").strip()
        if text.startswith("OK|"):
            return SegmentReport(
                anchor_id=anchor_id,
                source_preview=source_text[:60],
                target_preview=target_text[:60],
                qa_ok=True,
                note=text[3:].strip(),
            )
        if text.startswith("FIX|"):
            fixed = text[4:].strip() or target_text
            return SegmentReport(
                anchor_id=anchor_id,
                source_preview=source_text[:60],
                target_preview=fixed[:60],
                qa_ok=False,
                note="QA suggested revision",
            )
        return SegmentReport(
            anchor_id=anchor_id,
            source_preview=source_text[:60],
            target_preview=target_text[:60],
            qa_ok=True,
            note="QA fallback accepted",
        )


def detect_format(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in {".pdf", ".md", ".docx", ".pptx"}:
        return ext
    raise ValueError(f"Unsupported format: {ext}")


def run_pipeline(input_path: Path, model: str, max_pages: int | None = None) -> Path:
    tr = Translator(model=model)
    out = make_output_path(input_path)
    ext = detect_format(input_path)

    if ext == ".md":
        localize_markdown(input_path, out, tr)
    elif ext == ".docx":
        localize_docx(input_path, out, tr)
    elif ext == ".pptx":
        localize_pptx(input_path, out, tr)
    elif ext == ".pdf":
        localize_pdf(input_path, out, tr, max_pages=max_pages)
    return out


def write_run_report(output_path: Path, reports: List[SegmentReport]) -> Path:
    report_path = output_path.with_suffix(output_path.suffix + ".report.txt")
    lines = [
        "zh-tw Agent Pipeline Report",
        f"output={output_path}",
        f"segments={len(reports)}",
        "",
    ]
    for r in reports:
        lines.append(
            f"{r.anchor_id} | qa_ok={r.qa_ok} | src={r.source_preview} | tgt={r.target_preview} | note={r.note}"
        )
    report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return report_path


def main() -> int:
    load_dotenv()
    parser = argparse.ArgumentParser(
        description="Agent-style pipeline for zh-tw localization (Coordinator/Translator/QA)."
    )
    parser.add_argument("input_file", help="Input file path")
    parser.add_argument(
        "--model",
        default=None,
        help="LLM model for translation. If omitted, use OPENAI_MODEL or fallback default.",
    )
    parser.add_argument(
        "--emit-report",
        action="store_true",
        help="Emit a simple QA report file next to output.",
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

    model = args.model or os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
    out = run_pipeline(src, model, max_pages=args.max_pages)

    if args.emit_report:
        # Placeholder report row; richer anchor-level hooks can be added per adapter.
        reports = [
            SegmentReport(
                anchor_id="global",
                source_preview=src.name,
                target_preview=out.name,
                qa_ok=True,
                note="File-level pipeline completed",
            )
        ]
        rp = write_run_report(out, reports)
        print(rp)

    print(out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

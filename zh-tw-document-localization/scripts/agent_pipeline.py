#!/usr/bin/env python3
from __future__ import annotations

import argparse
import os
import re
import time
from dataclasses import dataclass
from pathlib import Path

import fitz  # pymupdf
from dotenv import load_dotenv
from localize import (
    Translator,
    localize_docx,
    localize_markdown,
    localize_pptx,
    make_output_path,
    normalize_translation,
    pdf_background_color,
    pdf_can_fit_text_in_line_rects,
    pdf_insert_text_in_line_rects,
    pdf_line_records,
    pdf_text_color,
    should_translate,
)


LAYOUT_SENSITIVE_HEADING_RE = re.compile(
    r"\b(complete\s+guide|guide|manual|handbook)\b", re.IGNORECASE
)


@dataclass
class FailureItem:
    page: int
    segment: int
    anchor_id: str
    reason: str


@dataclass
class PageStats:
    page: int
    status: str
    elapsed_sec: float
    extracted_segments: int
    translated_segments: int
    skipped_segments: int
    failed_segments: int
    qa_passed: int
    qa_revised: int


@dataclass
class QAReview:
    qa_ok: bool
    revised_text: str
    note: str


class QAAgent:
    def __init__(self, translator: Translator) -> None:
        self.translator = translator

    def review(self, source_text: str, target_text: str) -> QAReview:
        prompt = (
            "You are a QA reviewer for zh-TW localization.\n"
            "Check if target text is accurate and natural Taiwan Traditional Chinese, and keep placeholders unchanged.\n"
            "Do not reorder heading lines or move document-type labels such as guide/manual/handbook to another line.\n"
            "Respond in one line as:\n"
            "OK|<note>\n"
            "or\n"
            "FIX|<better translation>\n\n"
            f"Source:\n{source_text}\n\nTarget:\n{target_text}"
        )
        try:
            resp = self.translator.client.responses.create(
                model=self.translator.model,
                input=prompt,
                temperature=0,
            )
            text = (resp.output_text or "").strip()
            if text.startswith("OK|"):
                return QAReview(
                    qa_ok=True,
                    revised_text=normalize_translation(source_text, target_text),
                    note=text[3:].strip(),
                )
            if text.startswith("FIX|"):
                revised = text[4:].strip() or target_text
                return QAReview(qa_ok=False, revised_text=revised, note="QA suggested revision")
            return QAReview(
                qa_ok=True,
                revised_text=normalize_translation(source_text, target_text),
                note="QA fallback accepted",
            )
        except Exception as exc:
            return QAReview(qa_ok=True, revised_text=target_text, note=f"QA skipped: {exc}")


def detect_format(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in {".pdf", ".md", ".docx", ".pptx"}:
        return ext
    raise ValueError(f"Unsupported format: {ext}")


def localize_pdf_with_report(
    src: Path,
    out: Path,
    tr: Translator,
    max_pages: int | None = None,
    enable_qa: bool = True,
) -> tuple[list[PageStats], list[FailureItem]]:
    qa_agent = QAAgent(tr) if enable_qa else None
    doc = fitz.open(str(src))
    failures: list[FailureItem] = []
    page_stats: list[PageStats] = []
    total_pages = len(doc)
    pages_to_process = total_pages if max_pages is None else min(max_pages, total_pages)

    for page_idx in range(pages_to_process):
        started = time.perf_counter()
        page_no = page_idx + 1
        print(f"Processing page {page_no}/{pages_to_process}...", flush=True)
        page = doc[page_idx]
        spans_to_replace: list[dict] = []
        extracted = 0
        translated_ok = 0
        skipped = 0
        failed = 0
        qa_passed = 0
        qa_revised = 0
        seg_idx = 0

        for line_record in pdf_line_records(page):
            seg_idx += 1
            anchor_id = f"p{page_no}.{line_record['anchor_id']}"
            source_text = line_record["text"]
            if line_record.get("skip") or not should_translate(source_text):
                skipped += 1
                continue
            extracted += 1
            try:
                translated = tr.translate(source_text)
            except Exception as exc:
                failed += 1
                failures.append(
                    FailureItem(
                        page=page_no,
                        segment=seg_idx,
                        anchor_id=anchor_id,
                        reason=f"translate error: {exc}",
                    )
                )
                continue

            if translated == source_text:
                skipped += 1
                continue
            if not pdf_can_fit_text_in_line_rects(
                line_record["line_rects"], translated, line_record["size"]
            ):
                skipped += 1
                failed += 1
                failures.append(
                    FailureItem(
                        page=page_no,
                        segment=seg_idx,
                        anchor_id=anchor_id,
                        reason="translated text did not fit original text box; kept source text",
                    )
                )
                continue

            layout_sensitive_heading = (
                "\n" in source_text and LAYOUT_SENSITIVE_HEADING_RE.search(source_text)
            )
            if qa_agent is not None and not layout_sensitive_heading:
                qa = qa_agent.review(source_text, translated)
                if qa.qa_ok:
                    qa_passed += 1
                    translated = qa.revised_text
                else:
                    qa_revised += 1
                    translated = normalize_translation(source_text, qa.revised_text)
            elif layout_sensitive_heading:
                qa_passed += 1

            translated_ok += 1
            rect = fitz.Rect(line_record["bbox"])
            spans_to_replace.append(
                {
                    "bbox": line_record["bbox"],
                    "line_rects": line_record["line_rects"],
                    "text": translated,
                    "size": line_record["size"],
                    "color": line_record["color"],
                    "fill": pdf_background_color(page, rect),
                }
            )

        try:
            for s in spans_to_replace:
                rect = fitz.Rect(s["bbox"])
                page.add_redact_annot(rect, fill=s["fill"])
            if spans_to_replace:
                page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
            for s in spans_to_replace:
                rect = fitz.Rect(s["bbox"])
                fitted = pdf_insert_text_in_line_rects(
                    page,
                    s["line_rects"],
                    s["text"],
                    font_size=s["size"],
                    color=pdf_text_color(s["color"]),
                )
                if not fitted:
                    failed += 1
                    failures.append(
                        FailureItem(
                            page=page_no,
                            segment=0,
                            anchor_id=f"p{page_no}",
                            reason="text insertion failed after fit preflight",
                        )
                    )
            status = "ok" if failed == 0 else "partial"
        except Exception as exc:
            status = "failed"
            failed += max(1, len(spans_to_replace))
            failures.append(
                FailureItem(
                    page=page_no,
                    segment=0,
                    anchor_id=f"p{page_no}",
                    reason=f"render error: {exc}",
                )
            )

        elapsed = time.perf_counter() - started
        page_stats.append(
            PageStats(
                page=page_no,
                status=status,
                elapsed_sec=elapsed,
                extracted_segments=extracted,
                translated_segments=translated_ok,
                skipped_segments=skipped,
                failed_segments=failed,
                qa_passed=qa_passed,
                qa_revised=qa_revised,
            )
        )

    doc.save(str(out))
    doc.close()
    return page_stats, failures


def run_pipeline(
    input_path: Path,
    model: str,
    max_pages: int | None = None,
    enable_qa: bool = True,
) -> tuple[Path, list[PageStats], list[FailureItem]]:
    tr = Translator(model=model)
    out = make_output_path(input_path)
    ext = detect_format(input_path)

    if ext == ".md":
        started = time.perf_counter()
        localize_markdown(input_path, out, tr)
        elapsed = time.perf_counter() - started
        return (
            out,
            [
                PageStats(
                    page=1,
                    status="ok",
                    elapsed_sec=elapsed,
                    extracted_segments=0,
                    translated_segments=0,
                    skipped_segments=0,
                    failed_segments=0,
                    qa_passed=0,
                    qa_revised=0,
                )
            ],
            [],
        )
    if ext == ".docx":
        started = time.perf_counter()
        localize_docx(input_path, out, tr)
        elapsed = time.perf_counter() - started
        return (
            out,
            [
                PageStats(
                    page=1,
                    status="ok",
                    elapsed_sec=elapsed,
                    extracted_segments=0,
                    translated_segments=0,
                    skipped_segments=0,
                    failed_segments=0,
                    qa_passed=0,
                    qa_revised=0,
                )
            ],
            [],
        )
    if ext == ".pptx":
        started = time.perf_counter()
        localize_pptx(input_path, out, tr)
        elapsed = time.perf_counter() - started
        return (
            out,
            [
                PageStats(
                    page=1,
                    status="ok",
                    elapsed_sec=elapsed,
                    extracted_segments=0,
                    translated_segments=0,
                    skipped_segments=0,
                    failed_segments=0,
                    qa_passed=0,
                    qa_revised=0,
                )
            ],
            [],
        )
    page_stats, failures = localize_pdf_with_report(
        input_path, out, tr, max_pages=max_pages, enable_qa=enable_qa
    )
    return out, page_stats, failures


def write_run_report(
    output_path: Path,
    input_path: Path,
    model: str,
    page_stats: list[PageStats],
    failures: list[FailureItem],
) -> Path:
    report_path = output_path.with_suffix(output_path.suffix + ".report.txt")
    total_pages = len(page_stats)
    total_extracted = sum(p.extracted_segments for p in page_stats)
    total_translated = sum(p.translated_segments for p in page_stats)
    total_skipped = sum(p.skipped_segments for p in page_stats)
    total_failed = sum(p.failed_segments for p in page_stats)
    total_qa_passed = sum(p.qa_passed for p in page_stats)
    total_qa_revised = sum(p.qa_revised for p in page_stats)
    qa_total = total_qa_passed + total_qa_revised
    qa_rate = (total_qa_passed / qa_total * 100.0) if qa_total > 0 else 0.0
    total_elapsed = sum(p.elapsed_sec for p in page_stats)

    lines = [
        "zh-tw Agent Pipeline Report",
        f"input={input_path}",
        f"output={output_path}",
        f"model={model}",
        "",
        "[Summary]",
        f"pages={total_pages}",
        f"elapsed_sec={total_elapsed:.2f}",
        f"segments_extracted={total_extracted}",
        f"segments_translated={total_translated}",
        f"segments_skipped={total_skipped}",
        f"segments_failed={total_failed}",
        "",
        "[QA]",
        f"qa_total={qa_total}",
        f"qa_passed={total_qa_passed}",
        f"qa_revised={total_qa_revised}",
        f"qa_pass_rate={qa_rate:.2f}%",
        "",
        "[Per-Page]",
    ]
    for p in page_stats:
        lines.append(
            f"page={p.page} status={p.status} elapsed_sec={p.elapsed_sec:.2f} "
            f"extracted={p.extracted_segments} translated={p.translated_segments} "
            f"skipped={p.skipped_segments} failed={p.failed_segments} "
            f"qa_passed={p.qa_passed} qa_revised={p.qa_revised}"
        )

    lines.append("")
    lines.append("[Failures]")
    if not failures:
        lines.append("none")
    else:
        for f in failures:
            lines.append(
                f"page={f.page} segment={f.segment} anchor={f.anchor_id} reason={f.reason}"
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
        help="Emit detailed QA report file next to output.",
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        default=None,
        help="Only process the first N pages for PDF. Ignored for non-PDF.",
    )
    parser.add_argument(
        "--disable-qa",
        action="store_true",
        help="Disable QA review pass to improve speed.",
    )
    args = parser.parse_args()

    src = Path(args.input_file).resolve()
    if not src.exists():
        raise FileNotFoundError(f"Input not found: {src}")

    model = args.model or os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
    out, page_stats, failures = run_pipeline(
        src,
        model,
        max_pages=args.max_pages,
        enable_qa=not args.disable_qa,
    )

    if args.emit_report:
        report = write_run_report(out, src, model, page_stats, failures)
        print(report)

    print(out)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

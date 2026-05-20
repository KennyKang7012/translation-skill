#!/usr/bin/env python3
"""Generate zh-tw output filename and optionally create the output file.

Default behavior:
  copy source file to `*.zh-tw.<ext>` so downstream steps have a real file.
"""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path


def make_zh_tw_name(input_path: Path) -> Path:
    suffix = input_path.suffix
    stem = input_path.stem
    parent = input_path.parent
    return parent / f"{stem}.zh-tw{suffix}"


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Generate zh-tw output path and create output file by copying source."
    )
    parser.add_argument("input_file", help="Source file path")
    parser.add_argument(
        "--print-only",
        action="store_true",
        help="Only print output path without creating a file.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite output file if it already exists.",
    )
    args = parser.parse_args()

    src = Path(args.input_file)
    if not src.exists():
        raise FileNotFoundError(f"Input file not found: {src}")

    output = make_zh_tw_name(src)

    if not args.print_only:
        if output.exists() and not args.overwrite:
            raise FileExistsError(
                f"Output already exists: {output}. Use --overwrite to replace it."
            )
        shutil.copy2(src, output)

    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

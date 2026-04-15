from __future__ import annotations

import argparse
from pathlib import Path
import sys
from typing import List, Optional

from mdconvert_app.service import SUPPORTED_EXTENSIONS, convert_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="amesmarkdown",
        description="Convert local Office files (.docx, .pptx, .xlsx) into Markdown.",
    )
    parser.add_argument("source", help="Input file or directory")
    parser.add_argument("output", help="Output .md file or output directory")
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    source = Path(args.source)
    output = Path(args.output)

    if not source.exists():
        parser.error(f"Source does not exist: {source}")

    if source.is_file() and source.suffix.lower() not in SUPPORTED_EXTENSIONS:
        parser.error(f"Unsupported file type: {source.suffix}")

    results = convert_path(source, output)
    if not results:
        print("No supported Office files were found.", file=sys.stderr)
        return 1

    for result in results:
        print(f"Converted {result.source} -> {result.destination}")
        for warning in result.warnings:
            print(f"  warning: {warning}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

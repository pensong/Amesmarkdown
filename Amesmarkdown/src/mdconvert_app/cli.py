from __future__ import annotations

import argparse
from pathlib import Path
import sys
from typing import List, Optional

from mdconvert_app.service import SUPPORTED_EXTENSIONS, convert_path


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="amesmarkdown",
        description="Convert supported Office, PDF, and Markdown files between Markdown and common document formats.",
    )
    parser.add_argument("source", help="Input file or directory")
    parser.add_argument("output", help="Output file or output directory")
    parser.add_argument(
        "--to",
        dest="target_format",
        choices=["md", "docx", "pptx", "xlsx", "pdf"],
        help="Target format when the output path is a directory. Required for Markdown input.",
    )
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

    target_format = args.target_format
    if source.is_file() and source.suffix.lower() == ".md" and output.suffix.lower() not in {".docx", ".pptx", ".xlsx", ".pdf"} and not target_format:
        parser.error("Markdown input requires --to when the output path is a directory.")

    try:
        results = convert_path(source, output, target_format=target_format)
    except ValueError as exc:
        parser.error(str(exc))

    if not results:
        print("No supported files were found for the requested conversion.", file=sys.stderr)
        return 1

    for result in results:
        print(f"Converted {result.source} -> {result.destination}")
        for warning in result.warnings:
            print(f"  warning: {warning}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

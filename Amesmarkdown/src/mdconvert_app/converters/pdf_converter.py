from __future__ import annotations

from pathlib import Path

from mdconvert_app.markdown_utils import normalize_whitespace
from mdconvert_app.models import ConversionResult


def convert_pdf(source: Path, destination: Path) -> ConversionResult:
    try:
        from pypdf import PdfReader
    except ModuleNotFoundError as exc:
        raise RuntimeError("PDF conversion requires the `pypdf` package. Install project dependencies and retry.") from exc

    reader = PdfReader(str(source))
    lines: list[str] = [f"# {source.stem}", ""]
    warnings: list[str] = []

    for index, page in enumerate(reader.pages, start=1):
        lines.append(f"## Page {index}")
        lines.append("")

        text = normalize_whitespace(page.extract_text() or "")
        if text:
            lines.append(text)
            lines.append("")
        else:
            warnings.append(f"Page {index} did not contain extractable text.")
            lines.append("_No extractable text found on this page._")
            lines.append("")

    markdown = "\n".join(_trim(lines)).strip() + "\n"
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(markdown, encoding="utf-8")
    return ConversionResult(source=source, destination=destination, markdown=markdown, warnings=warnings)


def _trim(lines: list[str]) -> list[str]:
    trimmed = list(lines)
    while trimmed and trimmed[-1] == "":
        trimmed.pop()
    return trimmed

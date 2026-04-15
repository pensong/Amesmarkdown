from __future__ import annotations

from pathlib import Path

from mdconvert_app.converters import convert_docx, convert_pptx, convert_xlsx
from mdconvert_app.markdown_utils import slugify
from mdconvert_app.models import ConversionResult

SUPPORTED_EXTENSIONS = {".docx", ".pptx", ".xlsx"}


def convert_path(source: Path, output_path: Path) -> list[ConversionResult]:
    source = source.expanduser().resolve()
    output_path = output_path.expanduser().resolve()

    if source.is_file():
        destination = _destination_for_file(source, output_path)
        return [_convert_file(source, destination)]

    results: list[ConversionResult] = []
    output_path.mkdir(parents=True, exist_ok=True)
    for candidate in sorted(source.rglob("*")):
        if candidate.is_file() and candidate.suffix.lower() in SUPPORTED_EXTENSIONS:
            relative_parent = candidate.relative_to(source).parent
            destination = output_path / relative_parent / f"{slugify(candidate.stem)}.md"
            results.append(_convert_file(candidate, destination))
    return results


def _convert_file(source: Path, destination: Path) -> ConversionResult:
    ext = source.suffix.lower()
    if ext == ".docx":
        return convert_docx(source, destination)
    if ext == ".pptx":
        return convert_pptx(source, destination)
    if ext == ".xlsx":
        return convert_xlsx(source, destination)
    raise ValueError(f"Unsupported file type: {source}")


def _destination_for_file(source: Path, output_path: Path) -> Path:
    if output_path.suffix.lower() == ".md":
        return output_path
    return output_path / f"{slugify(source.stem)}.md"

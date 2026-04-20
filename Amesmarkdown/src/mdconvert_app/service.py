from __future__ import annotations

from pathlib import Path

from mdconvert_app.markdown_utils import slugify
from mdconvert_app.models import ConversionResult

MARKDOWN_EXTENSION = ".md"
MARKDOWN_EXPORT_EXTENSIONS = {".docx", ".pdf", ".pptx", ".xlsx"}
MARKDOWN_IMPORT_EXTENSIONS = {".docx", ".pdf", ".pptx", ".xlsx"}
SUPPORTED_EXTENSIONS = MARKDOWN_IMPORT_EXTENSIONS | {MARKDOWN_EXTENSION}


def convert_path(source: Path, output_path: Path, target_format: str | None = None) -> list[ConversionResult]:
    source = source.expanduser().resolve()
    output_path = output_path.expanduser().resolve()
    normalized_target = _normalize_target_format(target_format)

    if source.is_file():
        destination = _destination_for_file(source, output_path, normalized_target)
        _validate_conversion_pair(source, destination)
        return [_convert_file(source, destination)]

    results: list[ConversionResult] = []
    output_path.mkdir(parents=True, exist_ok=True)
    supported_sources = _directory_source_extensions(normalized_target)
    for candidate in sorted(source.rglob("*")):
        if candidate.is_file() and candidate.suffix.lower() in supported_sources:
            relative_parent = candidate.relative_to(source).parent
            destination = output_path / relative_parent / f"{slugify(candidate.stem)}{_target_extension_for(candidate, normalized_target)}"
            results.append(_convert_file(candidate, destination))
    return results


def _convert_file(source: Path, destination: Path) -> ConversionResult:
    ext = source.suffix.lower()
    if ext == MARKDOWN_EXTENSION:
        from mdconvert_app.converters.markdown_export import convert_markdown

        return convert_markdown(source, destination)
    if ext == ".docx":
        from mdconvert_app.converters.docx_converter import convert_docx

        return convert_docx(source, destination)
    if ext == ".pdf":
        from mdconvert_app.converters.pdf_converter import convert_pdf

        return convert_pdf(source, destination)
    if ext == ".pptx":
        from mdconvert_app.converters.pptx_converter import convert_pptx

        return convert_pptx(source, destination)
    if ext == ".xlsx":
        from mdconvert_app.converters.xlsx_converter import convert_xlsx

        return convert_xlsx(source, destination)
    raise ValueError(f"Unsupported file type: {source}")


def _destination_for_file(source: Path, output_path: Path, target_format: str | None) -> Path:
    if output_path.suffix.lower() in (MARKDOWN_EXPORT_EXTENSIONS | {MARKDOWN_EXTENSION}):
        return output_path
    return output_path / f"{slugify(source.stem)}{_target_extension_for(source, target_format)}"


def _target_extension_for(source: Path, target_format: str | None) -> str:
    source_ext = source.suffix.lower()
    if source_ext == MARKDOWN_EXTENSION:
        if target_format in MARKDOWN_EXPORT_EXTENSIONS:
            return target_format
        raise ValueError("Markdown input requires a target format of .docx, .pptx, .xlsx, or .pdf.")
    if target_format and target_format != MARKDOWN_EXTENSION:
        raise ValueError("Word, PowerPoint, Excel, and PDF inputs can only be converted to Markdown (.md).")
    return MARKDOWN_EXTENSION


def _directory_source_extensions(target_format: str | None) -> set[str]:
    if target_format in MARKDOWN_EXPORT_EXTENSIONS:
        return {MARKDOWN_EXTENSION}
    return MARKDOWN_IMPORT_EXTENSIONS


def _normalize_target_format(target_format: str | None) -> str | None:
    if not target_format:
        return None
    normalized = target_format.lower()
    if not normalized.startswith("."):
        normalized = f".{normalized}"
    if normalized not in (MARKDOWN_EXPORT_EXTENSIONS | {MARKDOWN_EXTENSION}):
        raise ValueError(f"Unsupported target format: {target_format}")
    return normalized


def _validate_conversion_pair(source: Path, destination: Path) -> None:
    source_ext = source.suffix.lower()
    destination_ext = destination.suffix.lower()
    if source_ext == MARKDOWN_EXTENSION and destination_ext not in MARKDOWN_EXPORT_EXTENSIONS:
        raise ValueError("Markdown input can only be exported to .docx, .pptx, .xlsx, or .pdf.")
    if source_ext in MARKDOWN_IMPORT_EXTENSIONS and destination_ext != MARKDOWN_EXTENSION:
        raise ValueError("Word, PowerPoint, Excel, and PDF inputs can only be converted to Markdown (.md).")

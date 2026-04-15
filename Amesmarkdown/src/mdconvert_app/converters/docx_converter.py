from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.document import Document as DocumentType
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from mdconvert_app.assets import AssetManager
from mdconvert_app.markdown_utils import ensure_blank_line, markdown_table, normalize_whitespace
from mdconvert_app.models import ConversionResult


def convert_docx(source: Path, destination: Path) -> ConversionResult:
    document = Document(source)
    assets = AssetManager(destination)
    lines: list[str] = [f"# {source.stem}", ""]
    warnings: list[str] = []

    for block in iter_block_items(document):
        if isinstance(block, Paragraph):
            _append_paragraph(lines, block, document, assets)
        elif isinstance(block, Table):
            _append_table(lines, block)

    markdown = "\n".join(_trim(lines)).strip() + "\n"
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_text(markdown, encoding="utf-8")
    return ConversionResult(source=source, destination=destination, markdown=markdown, warnings=warnings)


def iter_block_items(parent: DocumentType):
    parent_elm = parent.element.body
    for child in parent_elm.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)


def _append_paragraph(
    lines: list[str],
    paragraph: Paragraph,
    document: DocumentType,
    assets: AssetManager,
) -> None:
    text = normalize_whitespace(paragraph.text)
    style_name = paragraph.style.name.lower() if paragraph.style and paragraph.style.name else ""

    if not text and not _paragraph_has_image(paragraph):
        return

    if style_name.startswith("heading"):
        try:
            level = max(1, min(6, int(style_name.split()[-1])))
        except (TypeError, ValueError):
            level = 2
        ensure_blank_line(lines)
        lines.append(f"{'#' * level} {text}")
        lines.append("")
    elif "list bullet" in style_name:
        lines.append(f"- {text}")
    elif "list number" in style_name:
        lines.append(f"1. {text}")
    else:
        if text:
            lines.append(text)
            lines.append("")

    for run in paragraph.runs:
        blips = run.element.xpath(".//a:blip")
        for blip in blips:
            r_id = blip.get(qn("r:embed"))
            if not r_id:
                continue
            image_part = document.part.related_parts[r_id]
            ext = Path(image_part.filename).suffix or ".png"
            relative = assets.save_bytes(image_part.blob, ext, "docx-image")
            lines.append(f"![{Path(relative).stem}]({relative})")
            lines.append("")


def _append_table(lines: list[str], table: Table) -> None:
    ensure_blank_line(lines)
    rows = []
    for row in table.rows:
        rows.append([normalize_whitespace(cell.text) for cell in row.cells])
    if rows:
        lines.append(markdown_table(rows))
        lines.append("")


def _paragraph_has_image(paragraph: Paragraph) -> bool:
    return bool(paragraph._element.xpath(".//a:blip"))


def _trim(lines: list[str]) -> list[str]:
    trimmed = list(lines)
    while trimmed and trimmed[-1] == "":
        trimmed.pop()
    return trimmed


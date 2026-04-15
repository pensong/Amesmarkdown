from __future__ import annotations

import re


def normalize_whitespace(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def slugify(value: str) -> str:
    cleaned = re.sub(r"[^a-zA-Z0-9]+", "-", value.strip().lower()).strip("-")
    return cleaned or "document"


def escape_cell(value: str) -> str:
    return (value or "").replace("|", r"\|").replace("\n", "<br>")


def markdown_table(rows: list[list[str]]) -> str:
    if not rows:
        return ""
    width = max(len(row) for row in rows)
    normalized = [row + [""] * (width - len(row)) for row in rows]
    header = "| " + " | ".join(escape_cell(cell) for cell in normalized[0]) + " |"
    separator = "| " + " | ".join("---" for _ in normalized[0]) + " |"
    body = [
        "| " + " | ".join(escape_cell(cell) for cell in row) + " |"
        for row in normalized[1:]
    ]
    return "\n".join([header, separator, *body])


def ensure_blank_line(lines: list[str]) -> None:
    if lines and lines[-1] != "":
        lines.append("")


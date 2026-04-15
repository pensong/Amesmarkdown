from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class ConversionResult:
    source: Path
    destination: Path
    markdown: str
    warnings: list[str] = field(default_factory=list)

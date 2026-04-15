from __future__ import annotations

import hashlib
from pathlib import Path


class AssetManager:
    def __init__(self, output_markdown: Path) -> None:
        self.output_markdown = output_markdown
        self.asset_dir = output_markdown.parent / "assets" / output_markdown.stem
        self.asset_dir.mkdir(parents=True, exist_ok=True)
        self._written: dict[str, Path] = {}

    def save_bytes(self, payload: bytes, suffix: str, stem_hint: str) -> str:
        key = hashlib.sha256(payload).hexdigest()
        if key not in self._written:
            suffix = suffix if suffix.startswith(".") else f".{suffix}"
            filename = f"{stem_hint}-{key[:10]}{suffix}"
            path = self.asset_dir / filename
            path.write_bytes(payload)
            self._written[key] = path
        return self._relative_path(self._written[key])

    def _relative_path(self, path: Path) -> str:
        return path.relative_to(self.output_markdown.parent).as_posix()


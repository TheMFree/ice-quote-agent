"""Plain text attachments (.txt, .md, .csv)."""
from __future__ import annotations
from pathlib import Path


def parse(path: Path, out) -> None:
    try:
        txt = path.read_text(encoding="utf-8", errors="replace")
    except Exception:
        txt = path.read_bytes().decode("latin-1", errors="replace")
    if txt.strip():
        out.add_text(f"[Text file {path.name}]\n{txt}", source_label=path.name)

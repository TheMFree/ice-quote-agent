"""Image attachments — downsize and pass as image blocks to Claude Vision."""
from __future__ import annotations
import io
from pathlib import Path

MAX_DIM = 2000


def parse(path: Path, out) -> None:
    try:
        from PIL import Image
    except ImportError:
        with open(path, "rb") as f:
            raw = f.read()
        mt = _guess_media_type(path)
        out.add_image(raw, mt, source_label=path.name)
        return

    img = Image.open(path)
    img = img.convert("RGB") if img.mode not in ("RGB", "RGBA") else img
    w, h = img.size
    if max(w, h) > MAX_DIM:
        scale = MAX_DIM / max(w, h)
        img = img.resize((int(w * scale), int(h * scale)))

    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    out.add_image(buf.getvalue(), "image/jpeg", source_label=path.name)


def _guess_media_type(path: Path) -> str:
    ext = path.suffix.lower()
    return {
        ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
        ".png": "image/png", ".gif": "image/gif",
        ".webp": "image/webp",
    }.get(ext, "image/jpeg")

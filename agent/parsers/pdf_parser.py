"""PDF parser — tries text extraction first, falls back to rendering pages
as images for Claude Vision."""
from __future__ import annotations
import io
import logging
from pathlib import Path

log = logging.getLogger("ice_quote_agent")

MIN_TEXT_CHARS_PER_PAGE = 40


def parse(path: Path, out) -> None:
    try:
        import pdfplumber
    except ImportError:
        pdfplumber = None

    text_pages: list[str] = []
    image_fallback_pages: list[int] = []

    if pdfplumber is not None:
        with pdfplumber.open(str(path)) as pdf:
            for i, page in enumerate(pdf.pages):
                txt = (page.extract_text() or "").strip()
                if len(txt) < MIN_TEXT_CHARS_PER_PAGE:
                    image_fallback_pages.append(i)
                text_pages.append(txt)

    for i, txt in enumerate(text_pages):
        if txt:
            out.add_text(f"[PDF {path.name} — page {i + 1}]\n{txt}",
                         source_label=f"{path.name}#p{i+1}")

    if image_fallback_pages:
        try:
            import pypdf  # noqa: F401
            from pdf2image import convert_from_path  # type: ignore
            images = convert_from_path(str(path), dpi=150)
            for i in image_fallback_pages:
                if i < len(images):
                    buf = io.BytesIO()
                    images[i].save(buf, format="PNG")
                    out.add_image(buf.getvalue(), "image/png",
                                  source_label=f"{path.name}#p{i+1}")
        except Exception as e:
            log.warning("Could not render scanned PDF pages (%s): %s",
                        path.name, e)
            out.add_text(
                f"[PDF {path.name}: {len(image_fallback_pages)} page(s) appear scanned "
                f"but image rendering is unavailable on this host.]",
                source_label=path.name,
            )

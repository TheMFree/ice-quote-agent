"""Dispatch attachments to the right parser based on extension / MIME."""
from __future__ import annotations
import logging
import mimetypes
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Literal

from . import pdf_parser, docx_parser, image_parser, text_parser, audio_parser

log = logging.getLogger("ice_quote_agent")

BlockType = Literal["text", "image"]


@dataclass
class ContentBlock:
    kind: BlockType
    text: str = ""
    image_bytes: bytes = b""
    image_media_type: str = ""
    source_label: str = ""


@dataclass
class ParsedContent:
    blocks: List[ContentBlock] = field(default_factory=list)

    def add_text(self, text: str, source_label: str = "") -> None:
        if text and text.strip():
            self.blocks.append(ContentBlock(kind="text", text=text, source_label=source_label))

    def add_image(self, image_bytes: bytes, media_type: str, source_label: str = "") -> None:
        self.blocks.append(ContentBlock(
            kind="image", image_bytes=image_bytes,
            image_media_type=media_type, source_label=source_label
        ))

    def merge(self, other: "ParsedContent") -> None:
        self.blocks.extend(other.blocks)


_EXT_MAP = {
    ".pdf": "pdf", ".docx": "docx", ".doc": "docx",
    ".txt": "text", ".md": "text", ".csv": "text",
    ".jpg": "image", ".jpeg": "image", ".png": "image",
    ".gif": "image", ".webp": "image", ".bmp": "image",
    ".tif": "image", ".tiff": "image",
    # Voice-memo / audio formats accepted by Whisper.
    ".m4a": "audio", ".mp3": "audio", ".wav": "audio",
    ".ogg": "audio", ".oga": "audio", ".opus": "audio",
    ".webm": "audio", ".flac": "audio", ".amr": "audio",
    ".3gp": "audio", ".3gpp": "audio", ".aac": "audio",
    ".mp4": "audio",  # iOS sometimes wraps voice memos in mp4 container
}


def _kind_for(path: Path) -> str:
    ext = path.suffix.lower()
    if ext in _EXT_MAP:
        return _EXT_MAP[ext]
    mime, _ = mimetypes.guess_type(str(path))
    if mime:
        if mime.startswith("image/"):
            return "image"
        if mime.startswith("audio/"):
            return "audio"
        if mime == "application/pdf":
            return "pdf"
        if "word" in mime:
            return "docx"
        if mime.startswith("text/"):
            return "text"
    return "unknown"


def parse_attachment(path: Path) -> ParsedContent:
    kind = _kind_for(path)
    log.info("Parsing attachment %s as %s", path.name, kind)
    out = ParsedContent()
    try:
        if kind == "pdf":
            pdf_parser.parse(path, out)
        elif kind == "docx":
            docx_parser.parse(path, out)
        elif kind == "image":
            image_parser.parse(path, out)
        elif kind == "text":
            text_parser.parse(path, out)
        elif kind == "audio":
            audio_parser.parse(path, out)
        else:
            log.warning("Skipping unsupported attachment: %s", path.name)
    except Exception as e:
        log.exception("Failed to parse %s: %s", path.name, e)
        out.add_text(f"[Failed to parse attachment {path.name}: {e}]",
                     source_label=path.name)
    return out

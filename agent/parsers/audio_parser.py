"""Whisper audio transcription parser.

Transcribes voice-note attachments via the OpenAI Whisper API and feeds
the transcript into ParsedContent so the downstream extractor can pull
quote fields out of spoken job descriptions.
"""
from __future__ import annotations
import logging
import os
from pathlib import Path

log = logging.getLogger("ice_quote_agent")

# Whisper API hard cap is 25 MB per request.
MAX_BYTES = 25 * 1024 * 1024

# Domain prompt biases the model toward our jargon so it spells things
# like "RGS conduit", "Solar Mars 100", "fuel gas skid" correctly.
DOMAIN_PROMPT = (
    "ICE Contractors field voice memo. Industry: oil & gas, pipelines, "
    "instrumentation and electrical. Common terms: RGS conduit, EMT, "
    "PVC coated conduit, wellhead, fuel gas skid, MCC, VFD, PLC, "
    "Solar Mars 100 turbine, gas compressor, filter separator, generator, "
    "junction box, cable tray, NEC 2023, NFPA 70, T&M, lump sum, "
    "mobilization, change order, demobilization."
)


def parse(path: Path, out) -> None:
    """Transcribe an audio file and add the resulting text to `out`."""
    api_key = os.getenv("OPENAI_API_KEY", "").strip()
    if not api_key:
        log.warning("OPENAI_API_KEY not set; skipping audio attachment %s",
                    path.name)
        out.add_text(
            f"[Audio attachment {path.name} not transcribed: "
            "OPENAI_API_KEY missing]",
            source_label=path.name,
        )
        return

    size = path.stat().st_size
    if size > MAX_BYTES:
        log.warning("Audio %s exceeds Whisper 25MB limit (%d bytes); skipping",
                    path.name, size)
        out.add_text(
            f"[Audio attachment {path.name} too large to transcribe "
            f"({size} bytes; Whisper limit is 25MB)]",
            source_label=path.name,
        )
        return

    try:
        # Lazy import so the module only loads if audio is actually present.
        from openai import OpenAI
    except ImportError:
        log.error("openai package not installed; cannot transcribe %s",
                  path.name)
        out.add_text(
            f"[Audio attachment {path.name} not transcribed: "
            "openai package missing]",
            source_label=path.name,
        )
        return

    log.info("Transcribing audio %s (%d bytes) via Whisper", path.name, size)
    try:
        client = OpenAI(api_key=api_key)
        with open(path, "rb") as f:
            result = client.audio.transcriptions.create(
                model="whisper-1",
                file=f,
                response_format="text",
                prompt=DOMAIN_PROMPT,
            )
    except Exception as e:
        log.exception("Whisper transcription failed for %s", path.name)
        out.add_text(
            f"[Audio attachment {path.name} transcription failed: {e}]",
            source_label=path.name,
        )
        return

    transcript = (
        result.strip() if isinstance(result, str) else str(result).strip()
    )
    if not transcript:
        log.warning("Whisper returned empty transcript for %s", path.name)
        out.add_text(
            f"[Audio attachment {path.name} produced empty transcript]",
            source_label=path.name,
        )
        return

    log.info("Transcribed %s: %d chars", path.name, len(transcript))
    out.add_text(
        f"=== Voice memo: {path.name} ===\n{transcript}",
        source_label=f"voice:{path.name}",
    )

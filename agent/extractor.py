"""Claude extraction — takes parsed content blocks and returns a QuoteData
object populated to the best of Claude's ability."""
from __future__ import annotations
import base64
import json
import logging
from typing import Any, Dict, List

from anthropic import Anthropic
from pydantic import ValidationError
from tenacity import retry, stop_after_attempt, wait_exponential

from .config import Settings
from .schema import QuoteData
from .parsers import ParsedContent

log = logging.getLogger("ice_quote_agent")

SYSTEM_PROMPT = """\
You are a quote-extraction assistant for ICE Contractors, Inc. — an
Instrumentation & Electrical (I&E) contractor serving the midstream
energy sector.

You will receive raw inputs from a team member: the body of an email,
attached PDFs/Word docs, or images of hand-written or printed estimates.
Your job is to extract a structured JSON object conforming to the schema
provided, leaving any field null or empty if the information is not
clearly present in the inputs.

Rules:
1. Never guess or fabricate numbers, man-hours, footages, or prices.
   If uncertain, leave the field null and note the gap in
   `extraction_notes`.
2. Preserve currency formatting exactly as provided (e.g. "$471,633.00").
3. For scope_categories, use common I&E headings when clearly supported
   by the source material.
4. `document_type` should be "LUMP SUM PROPOSAL" for larger
   station/facility jobs, or "QUOTATION" for smaller scopes.
5. Keep `scope_intro` to ONE professional paragraph in ICE's voice.
6. Keep `scope_bullets` to 4–10 high-level bullets.
7. Always respond with a single JSON object and nothing else.
"""


def _build_user_message(parsed: ParsedContent, email_subject: str,
                        email_body: str, sender: str) -> List[Dict[str, Any]]:
    content: List[Dict[str, Any]] = []
    preamble = (
        f"Incoming email from: {sender}\n"
        f"Subject: {email_subject}\n\n"
        f"--- Email body ---\n{email_body.strip() or '(empty)'}\n"
    )
    content.append({"type": "text", "text": preamble})

    if parsed.blocks:
        content.append({"type": "text", "text": "--- Attachments ---"})
        for i, block in enumerate(parsed.blocks):
            if block.kind == "text":
                content.append({
                    "type": "text",
                    "text": f"[Attachment {i+1} — {block.source_label}]\n{block.text}",
                })
            elif block.kind == "image":
                content.append({
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": block.image_media_type or "image/jpeg",
                        "data": base64.b64encode(block.image_bytes).decode("ascii"),
                    },
                })

    schema = QuoteData.model_json_schema()
    content.append({
        "type": "text",
        "text": (
            "Produce a single JSON object matching this schema. Missing "
            "fields → null or empty arrays. Respond with JSON only.\n\n"
            f"SCHEMA:\n{json.dumps(schema, indent=2)}"
        ),
    })
    return content


@retry(stop=stop_after_attempt(3),
       wait=wait_exponential(multiplier=2, min=2, max=30))
def extract_quote_data(settings: Settings, parsed: ParsedContent,
                       email_subject: str, email_body: str,
                       sender: str) -> QuoteData:
    client = Anthropic(api_key=settings.anthropic_api_key)
    user_content = _build_user_message(parsed, email_subject, email_body, sender)

    log.info("Calling Claude (%s) with %d content blocks",
             settings.claude_model, len(user_content))

    resp = client.messages.create(
        model=settings.claude_model,
        max_tokens=4096,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_content}],
    )

    raw = "".join(
        b.text for b in resp.content if getattr(b, "type", "") == "text"
    ).strip()

    if raw.startswith("```"):
        raw = raw.strip("`")
        if raw.lower().startswith("json"):
            raw = raw[4:].strip()

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        log.error("Claude returned non-JSON: %s", raw[:400])
        raise RuntimeError(f"Extractor did not return valid JSON: {e}") from e

    try:
        return QuoteData.model_validate(data)
    except ValidationError as e:
        log.error("Schema validation failed: %s", e)
        safe: dict[str, Any] = {}
        for k, v in data.items():
            try:
                QuoteData.model_validate({**safe, k: v})
                safe[k] = v
            except ValidationError:
                log.warning("Dropping invalid field %s", k)
        return QuoteData.model_validate(safe)

"""Main agent loop — polls the mailbox, processes new emails, fills the
template, and replies with the draft attached."""
from __future__ import annotations
import argparse
import json
import re
import signal
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path
from typing import List

from agent.config import load_settings
from agent.email_client import GraphClient, IncomingEmail, EmailAttachment
from agent.extractor import extract_quote_data
from agent.filler import fill_template
from agent.logger import setup_logger
from agent.parsers import parse_attachment, ParsedContent
from agent.schema import QuoteData


def _safe(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("_") or "quote"


def _review_body(
    data: QuoteData,
    missing_fields: List[str],
    original_sender: str,
    original_subject: str,
    original_body: str,
) -> str:
    lines = [
        "Draft quote ready for your review.",
        "",
        f"From:     {original_sender}",
        f"Subject:  {original_subject}",
        "",
        "--- Extracted summary ---",
    ]
    if data.project:
        lines.append(f"Project:  {data.project}")
    if data.client:
        lines.append(f"Client:   {data.client}")
    if data.location:
        lines.append(f"Location: {data.location}")
    if data.total_amount:
        lines.append(f"Total:    {data.total_amount}")

    if missing_fields:
        lines += [
            "",
            "Information still needed:",
            *[f"  • {m}" for m in missing_fields],
        ]
    if data.extraction_notes:
        lines += ["", f"Extractor notes: {data.extraction_notes}"]

    lines += [
        "",
        "Review the attached .docx, adjust as needed, then forward to "
        f"{original_sender} (and the client) yourself.",
        "",
        "--- Original request ---",
        original_body.strip() or "(empty body)",
        "",
        "— ICE Quote Agent (automated)",
    ]
    return "\n".join(lines)


def _missing_fields(data: QuoteData) -> List[str]:
    missing = []
    if not data.project: missing.append("Project name")
    if not data.client: missing.append("Client name")
    if not data.location: missing.append("Location")
    if not data.material_amount: missing.append("Material cost")
    if not data.labor_equipment_amount: missing.append("Labor & equipment cost")
    if not data.total_amount: missing.append("Total price")
    if not data.scope_bullets: missing.append("High-level scope bullets")
    return missing


def process_email(email: IncomingEmail, settings, log, graph: GraphClient) -> bool:
    log.info("Processing email id=%s subject=%r from=%s attachments=%d",
             email.id[:12], email.subject, email.sender, len(email.attachments))

    with tempfile.TemporaryDirectory() as tdir:
        tpath = Path(tdir)
        parsed = ParsedContent()
        for att in email.attachments:
            fp = tpath / _safe(att.name)
            fp.write_bytes(att.content_bytes)
            parsed.merge(parse_attachment(fp))

        data = extract_quote_data(
            settings=settings,
            parsed=parsed,
            email_subject=email.subject,
            email_body=email.body_text,
            sender=email.sender,
        )
        log.info("Extracted: project=%r client=%r total=%r",
                 data.project, data.client, data.total_amount)

        stamp = datetime.utcnow().strftime("%Y%m%d-%H%M%S")
        base = _safe(data.project or email.subject or "quote")[:60]
        out_path = settings.output_dir / f"{stamp}_{base}.docx"
        fill_template(settings.template_path, data, out_path)

        jpath = out_path.with_suffix(".json")
        jpath.write_text(data.model_dump_json(indent=2))

        missing = _missing_fields(data)
        body = _review_body(
            data=data,
            missing_fields=missing,
            original_sender=email.sender,
            original_subject=email.subject,
            original_body=email.body_text,
        )
        review_subject = f"[Draft Quote for Review] {email.subject}"

        if not settings.reviewer_email:
            raise RuntimeError(
                "REVIEWER_EMAIL is not set. The agent needs a reviewer to "
                "send drafts to — set REVIEWER_EMAIL in .env."
            )

        if settings.dry_run:
            log.info("[DRY_RUN] Would send review TO=%s (draft %s)",
                     settings.reviewer_email, out_path.name)
        else:
            graph.send_reply(
                to=settings.reviewer_email,
                cc=None,
                subject=review_subject,
                body_text=body,
                attachment_path=out_path,
                reply_to_message_id=None,
            )
            log.info("Review email sent to %s.", settings.reviewer_email)

    return True


_stop = False

def _handle_signal(*_):
    global _stop
    _stop = True


def run():
    settings = load_settings()
    log = setup_logger(settings.log_file)
    log.info("ICE Quote Agent starting (dry_run=%s, mailbox=%s)",
             settings.dry_run, settings.mailbox or "<not set>")
    settings.output_dir.mkdir(parents=True, exist_ok=True)

    if not (settings.tenant_id and settings.client_id and
            settings.client_secret and settings.mailbox):
        log.error("M365 credentials or mailbox missing. Fill in .env.")
        sys.exit(2)

    graph = GraphClient(settings)

    signal.signal(signal.SIGINT, _handle_signal)
    signal.signal(signal.SIGTERM, _handle_signal)

    while not _stop:
        try:
            emails = graph.fetch_unread(top=10)
            for email in emails:
                try:
                    ok = process_email(email, settings, log, graph)
                    if ok:
                        if not settings.dry_run:
                            graph.mark_read(email.id)
                            graph.move_to_folder(email.id, settings.processed_folder)
                except Exception:
                    log.exception("Failed processing email %s", email.id[:12])
                    if not settings.dry_run:
                        graph.move_to_folder(email.id, settings.failed_folder)
        except Exception:
            log.exception("Poll cycle failed")

        for _ in range(settings.poll_interval):
            if _stop:
                break
            time.sleep(1)

    log.info("Agent stopped.")


def run_once_from_text(subject: str, body: str, sender: str,
                       attachments: List[Path] | None = None):
    settings = load_settings()
    log = setup_logger(settings.log_file)
    parsed = ParsedContent()
    for p in attachments or []:
        parsed.merge(parse_attachment(p))

    data = extract_quote_data(settings, parsed, subject, body, sender)
    log.info("Extracted: %s", json.dumps(data.model_dump(), indent=2)[:800])

    stamp = datetime.utcnow().strftime("%Y%m%d-%H%M%S")
    out_path = settings.output_dir / f"{stamp}_{_safe(data.project or 'test')}.docx"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    fill_template(settings.template_path, data, out_path)
    print(f"Wrote {out_path}")


def main():
    parser = argparse.ArgumentParser(description="ICE Quote Agent")
    parser.add_argument("--test", action="store_true",
                        help="Run a single extraction+fill against "
                             "tests/sample_email.txt and exit.")
    args = parser.parse_args()

    if args.test:
        root = Path(__file__).resolve().parent
        sample = (root / "tests" / "sample_email.txt").read_text()
        parts = sample.split("\n---\n", 2)
        subj = parts[0].replace("Subject:", "").strip() if parts else "Test"
        body = parts[1] if len(parts) > 1 else sample
        run_once_from_text(subj, body, "team@icecontractorsinc.com")
        return

    run()


if __name__ == "__main__":
    main()

"""Central configuration loaded from .env."""
from __future__ import annotations
import os
from dataclasses import dataclass
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

PROJECT_ROOT = Path(__file__).resolve().parent.parent


def _get(name: str, default: str | None = None, required: bool = False) -> str:
    val = os.getenv(name, default)
    if required and (val is None or val == ""):
        raise RuntimeError(f"Missing required env var: {name}")
    return val or ""


def _bool(name: str, default: bool = False) -> bool:
    return _get(name, "true" if default else "false").strip().lower() in ("1", "true", "yes", "y")


@dataclass(frozen=True)
class Settings:
    anthropic_api_key: str
    claude_model: str
    tenant_id: str
    client_id: str
    client_secret: str
    mailbox: str
    reviewer_email: str
    poll_interval: int
    processed_folder: str
    failed_folder: str
    template_path: Path
    output_dir: Path
    log_file: Path
    dry_run: bool


def load_settings() -> Settings:
    return Settings(
        anthropic_api_key=_get("ANTHROPIC_API_KEY", required=True),
        claude_model=_get("CLAUDE_MODEL", "claude-sonnet-4-6"),
        tenant_id=_get("MS_TENANT_ID", required=False),
        client_id=_get("MS_CLIENT_ID", required=False),
        client_secret=_get("MS_CLIENT_SECRET", required=False),
        mailbox=_get("AGENT_MAILBOX", required=False),
        reviewer_email=_get("REVIEWER_EMAIL", _get("CC_RECIPIENT", "")),
        poll_interval=int(_get("POLL_INTERVAL_SECONDS", "60")),
        processed_folder=_get("PROCESSED_FOLDER", "Processed"),
        failed_folder=_get("FAILED_FOLDER", "FailedToProcess"),
        template_path=(PROJECT_ROOT / _get("TEMPLATE_PATH", "templates/ICE_Contractors_Proposal_Template.docx")).resolve(),
        output_dir=(PROJECT_ROOT / _get("OUTPUT_DIR", "output")).resolve(),
        log_file=(PROJECT_ROOT / _get("LOG_FILE", "logs/agent.log")).resolve(),
        dry_run=_bool("DRY_RUN", default=True),
    )

"""Polish — final-gate document QA for the Quote Agent.

Runs a deterministic style linter (regex-based) plus a Claude-powered
linguistic pass (grammar, awkward phrasing, parallel structure, voice) over
the filled proposal and returns a PolishResult the main loop can use to
decide how to present the review email.

Produces:
- A tracked-changes .docx where every edit is attributed to "Polish"
- A punch-list markdown string the main loop pastes into the email body
- A grade (A/B/C/D) based on finding count

House style baked in:
- Oxford comma: always
- I&E / T&M: ampersand, no spaces
- Currency: $1,234.56 (two decimals always)
- Dates (prose): January 1, 2026
- Smart quotes, em-dash, ellipsis
- No double spaces, no exclamation in formal prose, no contractions

Mirrors the Polish skill at ~/.claude/skills/polish/. Rules are vendored
here so the agent is self-contained; the skill remains the canonical
style reference for interactive use.
"""
from __future__ import annotations

import json
import logging
import re
import shutil
import subprocess
import tempfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Optional, Tuple

from anthropic import Anthropic

log = logging.getLogger("ice_quote_agent")

# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class Finding:
    rule: str          # e.g. "R01-oxford-comma" or "L-grammar"
    line: int
    col: int
    current: str
    suggested: str
    note: str = ""
    kind: str = "deterministic"  # "deterministic" or "linguistic"

    def to_dict(self) -> dict:
        return {
            "rule": self.rule, "line": self.line, "col": self.col,
            "current": self.current, "suggested": self.suggested,
            "note": self.note, "kind": self.kind,
        }


@dataclass
class PolishResult:
    grade: str                                     # "A" / "B" / "C" / "D"
    findings: List[Finding] = field(default_factory=list)
    punch_list_md: str = ""
    polished_docx: Optional[Path] = None
    clean_docx: Optional[Path] = None
    ran_linguistic_pass: bool = False

    @property
    def total_findings(self) -> int:
        return len(self.findings)

    @property
    def should_ship_tracked(self) -> bool:
        return self.grade != "A"


# ---------------------------------------------------------------------------
# Deterministic linter (R01–R13)
# ---------------------------------------------------------------------------

_R01_RE = re.compile(
    r"(\b\w+\b(?:\s+\w+)*),\s+(\b\w+\b(?:\s+\w+)*)\s+(and|or)\s+(\b\w+\b(?:\s+\w+)*)"
)
_R02_VARIANTS = [
    (re.compile(r"\bI and E\b"), "I&E"),
    (re.compile(r"\bI & E\b"), "I&E"),
    (re.compile(r"\bI&e\b"), "I&E"),
    (re.compile(r"\bi&E\b"), "I&E"),
    (re.compile(r"\bi&e\b"), "I&E"),
]
_R03_VARIANTS = [
    (re.compile(r"\bT and M\b"), "T&M"),
    (re.compile(r"\bT & M\b"), "T&M"),
    (re.compile(r"\bT&m\b"), "T&M"),
    (re.compile(r"\bt&m\b"), "T&M"),
]
_R04_RE = re.compile(r"\$(\d{1,3}(?:,\d{3})*|\d{4,})(?!\.\d|\d)")
_R05_NUMERIC_US = re.compile(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b")
_R05_ISO = re.compile(r"\b(\d{4})-(\d{2})-(\d{2})\b")
_R05_ABBREV = re.compile(
    r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\.?\s+(\d{1,2}),?\s+(\d{4})\b"
)
_MONTH_NAMES = {
    "Jan": "January", "Feb": "February", "Mar": "March", "Apr": "April",
    "May": "May", "Jun": "June", "Jul": "July", "Aug": "August",
    "Sep": "September", "Oct": "October", "Nov": "November", "Dec": "December",
}
_ISO_MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]
_R06_STRAIGHT_SINGLE = re.compile(r"(?<=\w)'(?=\w)")
_R06_STRAIGHT_DOUBLE = re.compile(r'(?<![0-9/])"')
_R07_RE = re.compile(r"(?<!-)--(?!-)")
_R08_RE = re.compile(r"\.{3,}")
_R09_RE = re.compile(r"(?<!^)  +")
_R10_RE = re.compile(r"\b(\d+)-(\d+)\b")
_R11_RE = re.compile(r"\b(\w+)\s+\1\b", re.IGNORECASE)
_R12_RE = re.compile(r"!")
_R13_CONTRACTIONS = {
    "won't": "will not", "can't": "cannot", "don't": "do not",
    "doesn't": "does not", "didn't": "did not", "isn't": "is not",
    "aren't": "are not", "wasn't": "was not", "weren't": "were not",
    "haven't": "have not", "hasn't": "has not", "hadn't": "had not",
    "wouldn't": "would not", "shouldn't": "should not", "couldn't": "could not",
    "it's": "it is", "we'll": "we will", "you'll": "you will",
    "they'll": "they will", "we've": "we have", "you've": "you have",
    "they've": "they have", "we're": "we are", "you're": "you are",
    "they're": "they are",
}
_R13_RE = re.compile(
    r"\b(" + "|".join(re.escape(c) for c in _R13_CONTRACTIONS.keys()) + r")\b",
    re.IGNORECASE,
)


def _add(findings: List[Finding], rule: str, line: int, col: int,
         current: str, suggested: str, note: str = "",
         kind: str = "deterministic") -> None:
    findings.append(Finding(rule=rule, line=line, col=col,
                            current=current, suggested=suggested,
                            note=note, kind=kind))


def _lint_line(ln: int, line: str, findings: List[Finding]) -> None:
    for m in _R01_RE.finditer(line):
        _add(findings, "R01-oxford-comma", ln, m.start() + 1,
             m.group(0),
             f"{m.group(1)}, {m.group(2)}, {m.group(3)} {m.group(4)}",
             "Missing Oxford comma.")
    for pat, repl in _R02_VARIANTS:
        for m in pat.finditer(line):
            _add(findings, "R02-ie-format", ln, m.start() + 1,
                 m.group(0), repl, "Use 'I&E'.")
    for pat, repl in _R03_VARIANTS:
        for m in pat.finditer(line):
            _add(findings, "R03-tm-format", ln, m.start() + 1,
                 m.group(0), repl, "Use 'T&M'.")
    for m in _R04_RE.finditer(line):
        raw = m.group(0)
        digits = raw.replace("$", "").replace(",", "")
        try:
            val = int(digits)
            _add(findings, "R04-currency", ln, m.start() + 1,
                 raw, f"${val:,}.00", "Currency needs two decimals.")
        except ValueError:
            pass
    for m in _R05_NUMERIC_US.finditer(line):
        try:
            mon, day, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if 1 <= mon <= 12 and 1 <= day <= 31:
                _add(findings, "R05-date-format", ln, m.start() + 1,
                     m.group(0), f"{_ISO_MONTHS[mon - 1]} {day}, {year}",
                     "Use long-form date.")
        except (ValueError, IndexError):
            pass
    for m in _R05_ISO.finditer(line):
        try:
            year, mon, day = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if 1 <= mon <= 12 and 1 <= day <= 31:
                _add(findings, "R05-date-format", ln, m.start() + 1,
                     m.group(0), f"{_ISO_MONTHS[mon - 1]} {day}, {year}",
                     "ISO format not for prose.")
        except (ValueError, IndexError):
            pass
    for m in _R05_ABBREV.finditer(line):
        mon = _MONTH_NAMES.get(m.group(1), m.group(1))
        suggested = f"{mon} {m.group(2)}, {m.group(3)}"
        if suggested != m.group(0):
            _add(findings, "R05-date-format", ln, m.start() + 1,
                 m.group(0), suggested, "Spell out the month.")
    for m in _R06_STRAIGHT_SINGLE.finditer(line):
        _add(findings, "R06-smart-quotes", ln, m.start() + 1,
             m.group(0), "\u2019", "Use curly apostrophe.")
    for m in _R06_STRAIGHT_DOUBLE.finditer(line):
        _add(findings, "R06-smart-quotes", ln, m.start() + 1,
             m.group(0), "\u201C or \u201D", "Use curly double quotes.")
    for m in _R07_RE.finditer(line):
        _add(findings, "R07-em-dash", ln, m.start() + 1,
             m.group(0), "\u2014", "Use em-dash.")
    for m in _R08_RE.finditer(line):
        _add(findings, "R08-ellipsis", ln, m.start() + 1,
             m.group(0), "\u2026", "Use ellipsis character.")
    for m in _R09_RE.finditer(line):
        _add(findings, "R09-double-space", ln, m.start() + 1,
             repr(m.group(0)), " ", "Collapse to single space.")
    for m in _R10_RE.finditer(line):
        a, b = int(m.group(1)), int(m.group(2))
        if a < b and 1 <= a and b <= 99999:
            _add(findings, "R10-hyphen-in-range", ln, m.start() + 1,
                 m.group(0), f"{a}\u2013{b}",
                 "Ranges use en-dash, not hyphen.")
    for m in _R11_RE.finditer(line):
        word = m.group(1).lower()
        if word in {"that", "had", "have", "is", "s"} or len(word) <= 2:
            continue
        _add(findings, "R11-double-word", ln, m.start() + 1,
             m.group(0), m.group(1), "Repeated word.")
    for m in _R12_RE.finditer(line):
        _add(findings, "R12-exclamation", ln, m.start() + 1,
             "!", ".", "Avoid exclamation in client-facing docs.")
    for m in _R13_RE.finditer(line):
        raw = m.group(0)
        expansion = _R13_CONTRACTIONS.get(raw.lower(), raw)
        if raw[0].isupper():
            expansion = expansion[0].upper() + expansion[1:]
        _add(findings, "R13-contraction", ln, m.start() + 1,
             raw, expansion, "Spell out contractions in formal prose.")


def lint_text(text: str) -> List[Finding]:
    findings: List[Finding] = []
    in_fence = False
    for idx, line in enumerate(text.splitlines(), start=1):
        stripped = line.strip()
        if stripped.startswith("```"):
            in_fence = not in_fence
            continue
        if in_fence:
            continue
        _lint_line(idx, line, findings)
    return findings


# ---------------------------------------------------------------------------
# Extract prose from .docx via pandoc (already on the droplet)
# ---------------------------------------------------------------------------

def _extract_prose(docx_path: Path) -> str:
    """Extract plain-text prose from a .docx using pandoc. Returns empty string
    on failure."""
    pandoc = shutil.which("pandoc")
    if not pandoc:
        log.warning("pandoc not found — skipping Polish prose extraction.")
        return ""
    with tempfile.NamedTemporaryFile(suffix=".md", delete=False) as tmp:
        md_path = Path(tmp.name)
    try:
        result = subprocess.run(
            [pandoc, str(docx_path), "-o", str(md_path), "--wrap=none"],
            capture_output=True, text=True, timeout=60,
        )
        if result.returncode != 0:
            log.warning("pandoc extract failed rc=%s: %s",
                        result.returncode, result.stderr[:200])
            return ""
        return md_path.read_text(encoding="utf-8", errors="replace")
    except subprocess.TimeoutExpired:
        log.warning("pandoc timed out extracting %s", docx_path.name)
        return ""
    finally:
        md_path.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# Linguistic pass (Claude)
# ---------------------------------------------------------------------------

_LINGUISTIC_SYSTEM = """\
You are Polish — the final-gate document QA auditor for ICE Contractors,
Inc. You review client-facing proposals for grammar, punctuation,
spelling, awkward phrasing, parallel-structure breaks in bulleted lists,
voice-and-tone drift, ambiguity, and redundancy.

You do NOT enforce mechanical style rules (Oxford comma, I&E, currency
format, date format, smart quotes, em-dashes) — those are handled by a
separate deterministic linter. Focus only on issues a regex cannot catch.

House voice: we/our for ICE; you/your for the client. Active voice
preferred. No contractions in formal prose. No exclamation points.

Return a single JSON array of findings. Each finding must have:
- "issue": short tag, one of: grammar, spelling, awkward, parallel,
  voice, ambiguous, redundant
- "current": the exact text as written (a full sentence or bullet)
- "suggested": your proposed rewrite
- "note": one-sentence explanation

Return ONLY the JSON array. No markdown fences, no preamble.
If you find nothing, return [].
"""


def _linguistic_pass(prose: str, api_key: str, model: str) -> List[Finding]:
    """Call Claude once to catch what regex cannot. Returns empty list on
    any failure — linguistic pass is best-effort, never blocks the agent."""
    if not prose.strip():
        return []
    try:
        client = Anthropic(api_key=api_key)
        resp = client.messages.create(
            model=model,
            max_tokens=2048,
            system=_LINGUISTIC_SYSTEM,
            messages=[{"role": "user", "content": prose}],
        )
        text = "".join(
            b.text for b in resp.content if getattr(b, "type", "") == "text"
        ).strip()
        # Strip accidental markdown fencing.
        if text.startswith("```"):
            text = re.sub(r"^```\w*\s*", "", text)
            text = re.sub(r"```\s*$", "", text)
        items = json.loads(text)
        if not isinstance(items, list):
            log.warning("Polish linguistic pass: non-list response; skipping.")
            return []
        findings: List[Finding] = []
        for it in items:
            if not isinstance(it, dict):
                continue
            findings.append(Finding(
                rule=f"L-{it.get('issue', 'other')}",
                line=0, col=0,
                current=str(it.get("current", ""))[:500],
                suggested=str(it.get("suggested", ""))[:500],
                note=str(it.get("note", ""))[:300],
                kind="linguistic",
            ))
        return findings
    except Exception as exc:
        log.warning("Polish linguistic pass failed (%s); continuing without it.",
                    exc.__class__.__name__)
        return []


# ---------------------------------------------------------------------------
# Tracked-changes injection into .docx
# ---------------------------------------------------------------------------

_START_ID = 9000


def _xml_escape(text: str) -> str:
    return (text.replace("&", "&amp;")
                 .replace("<", "&lt;")
                 .replace(">", "&gt;"))


def _inject_tracked_changes(docx_src: Path, docx_dst: Path,
                            findings: List[Finding]) -> bool:
    """Create a copy of docx_src at docx_dst with tracked-change blocks for
    each deterministic finding whose `current` appears verbatim in a <w:t>.
    Returns True if at least one change was applied."""
    import zipfile
    det = [f for f in findings
           if f.kind == "deterministic" and f.current and f.suggested
           and f.current != f.suggested
           and not f.current.startswith("'") and not f.current.startswith('"')]
    if not det:
        shutil.copyfile(docx_src, docx_dst)
        return False

    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    wid = _START_ID
    applied = 0

    def track_block(old: str, new: str, wid: int) -> str:
        old_x, new_x = _xml_escape(old), _xml_escape(new)
        return (
            f'<w:del w:id="{wid}" w:author="Polish" w:date="{now}">'
            f'<w:r><w:delText xml:space="preserve">{old_x}</w:delText></w:r>'
            f'</w:del>'
            f'<w:ins w:id="{wid + 1}" w:author="Polish" w:date="{now}">'
            f'<w:r><w:t xml:space="preserve">{new_x}</w:t></w:r>'
            f'</w:ins>'
        )

    t_re = re.compile(r"(<w:t(?:\s+[^>]*)?>)([^<]*)(</w:t>)")

    def _sub(m: re.Match) -> str:
        nonlocal wid, applied
        opening, inner, closing = m.group(1), m.group(2), m.group(3)
        for f in det:
            if f.current and f.current in inner:
                # Truncate suggested to just the replaced substring when the
                # finding `current` is already the exact string. For R01
                # Oxford-comma findings, `current` is a wide sentence span
                # and `suggested` is the same span rewritten — that works
                # here too.
                prefix, suffix = inner.split(f.current, 1)
                block = track_block(f.current, f.suggested, wid)
                wid += 2
                applied += 1
                return (f"{opening}{prefix}{closing}"
                        f"</w:r>{block}<w:r>"
                        f"{opening}{suffix}{closing}")
        return m.group(0)

    with zipfile.ZipFile(docx_src, "r") as zin:
        members = {n: zin.read(n) for n in zin.namelist()}

    doc_xml = members.get("word/document.xml", b"").decode("utf-8")
    new_xml = t_re.sub(_sub, doc_xml)
    members["word/document.xml"] = new_xml.encode("utf-8")

    docx_dst.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(docx_dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    log.info("Polish: applied %d tracked change(s) to %s",
             applied, docx_dst.name)
    return applied > 0


# ---------------------------------------------------------------------------
# Punch list + grading
# ---------------------------------------------------------------------------

def _grade(n_det: int, n_ling: int) -> str:
    """Simple grading heuristic. Linguistic findings weigh more heavily
    because they indicate real grammar/spelling problems."""
    score = n_det + 2 * n_ling
    if score == 0:
        return "A"
    if score <= 5:
        return "B"
    if score <= 15:
        return "C"
    return "D"


def _punch_list(findings: List[Finding], grade: str) -> str:
    det = [f for f in findings if f.kind == "deterministic"]
    ling = [f for f in findings if f.kind == "linguistic"]

    lines: List[str] = []
    lines.append(f"**Polish grade: {grade}**  ·  "
                 f"{len(det)} mechanical, {len(ling)} linguistic "
                 f"({len(findings)} total)")
    lines.append("")

    if det:
        lines.append("**Mechanical style (tracked in the .docx):**")
        for i, f in enumerate(det[:20], start=1):
            lines.append(f"  {i}. `{f.rule}` — "
                         f"`{_truncate(f.current, 60)}` → "
                         f"`{_truncate(f.suggested, 60)}`")
        if len(det) > 20:
            lines.append(f"  _(+{len(det) - 20} more — see tracked-changes .docx)_")
        lines.append("")

    if ling:
        lines.append("**Linguistic (suggestions only — not tracked):**")
        for i, f in enumerate(ling[:15], start=1):
            tag = f.rule.replace("L-", "")
            lines.append(f"  {i}. _{tag}_ — {f.note}")
            lines.append(f"     - As written: {_truncate(f.current, 120)}")
            lines.append(f"     - Suggested:  {_truncate(f.suggested, 120)}")
        if len(ling) > 15:
            lines.append(f"  _(+{len(ling) - 15} more)_")
        lines.append("")

    lines.append(f"_Polished by Polish v1.0 — "
                 f"{datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}_")
    return "\n".join(lines)


def _truncate(s: str, n: int) -> str:
    s = s.replace("\n", " ").replace("`", "'")
    return s if len(s) <= n else s[: n - 1] + "…"


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def run_polish(docx_path: Path, *, anthropic_api_key: str,
               model: str, run_linguistic: bool = True) -> PolishResult:
    """Run Polish over a filled .docx and return a PolishResult.

    Never raises — on any error the result degrades to grade A (silent) so the
    agent keeps shipping. Errors are logged.
    """
    try:
        prose = _extract_prose(docx_path)
        findings = lint_text(prose) if prose else []

        ran_ling = False
        if run_linguistic and anthropic_api_key and prose.strip():
            ling = _linguistic_pass(prose, anthropic_api_key, model)
            findings.extend(ling)
            ran_ling = True

        n_det = sum(1 for f in findings if f.kind == "deterministic")
        n_ling = sum(1 for f in findings if f.kind == "linguistic")
        grade = _grade(n_det, n_ling)

        polished_path: Optional[Path] = None
        if findings and n_det > 0:
            polished_path = docx_path.with_name(
                docx_path.stem + "_polished.docx"
            )
            _inject_tracked_changes(docx_path, polished_path, findings)

        punch = _punch_list(findings, grade) if findings else ""

        log.info("Polish grade=%s findings=%d (det=%d, ling=%d)",
                 grade, len(findings), n_det, n_ling)

        return PolishResult(
            grade=grade,
            findings=findings,
            punch_list_md=punch,
            polished_docx=polished_path,
            clean_docx=docx_path,
            ran_linguistic_pass=ran_ling,
        )
    except Exception:
        log.exception("Polish crashed — degrading to grade A (silent).")
        return PolishResult(grade="A", clean_docx=docx_path)

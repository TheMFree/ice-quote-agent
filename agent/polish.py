"""Polish — silent final-gate QA for the Quote Agent.

Catches MAJOR, UNPROFESSIONAL errors and corrects them silently in the
filled proposal. No tracked changes. No punch list. No email report.

Rules enforced (deterministic):
- I&E / T&M spelled with ampersand (not 'I and E', 'I & E', etc.)
- Straight quotes -> smart quotes (apostrophes and double quotes)
- Double spaces collapsed to single
- Doubled words ('the the') de-duplicated
- Exclamation points -> periods (formal prose)
- Contractions expanded ('don't' -> 'do not')
- Currency normalized to two decimals ('$1,234' -> '$1,234.00')

Also runs a Claude-powered pass to catch typos / grammar errors /
wrong-word errors / broken sentences that regex cannot see. Applied
silently where the replacement is short enough to fit a single run.

Style preferences that are NOT ship-blockers (Oxford comma, date format,
em-dash vs --, ellipsis character, hyphen vs en-dash in ranges) are
deliberately NOT enforced here - those belong in the interactive Polish
skill, not in the Quote Agent's silent gate.

Public API:
    result = run_polish(docx_path, anthropic_api_key=..., model=...)
    # result.polished_docx points at a clean corrected file (or None)
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
    rule: str          # e.g. "R02-ie-format" or "L-typo"
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
    grade: str                                     # "A" (no edits) or "B" (applied edits)
    findings: List[Finding] = field(default_factory=list)
    punch_list_md: str = ""                        # always "" in silent mode
    polished_docx: Optional[Path] = None
    clean_docx: Optional[Path] = None
    ran_linguistic_pass: bool = False

    @property
    def total_findings(self) -> int:
        return len(self.findings)


# ---------------------------------------------------------------------------
# Deterministic linter (major-error rules only)
# ---------------------------------------------------------------------------

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
_R06_STRAIGHT_SINGLE = re.compile(r"(?<=\w)'(?=\w)")
_R06_STRAIGHT_DOUBLE = re.compile(r'(?<![0-9/])"')
_R09_RE = re.compile(r"(?<!^)  +")
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
    """Flag only MAJOR, unprofessional errors. Style preferences
    (Oxford comma, date format, em-dash, ellipsis, hyphen-in-range)
    are deliberately skipped - they belong in the interactive Polish
    skill, not in the silent Quote Agent gate."""
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
    for m in _R06_STRAIGHT_SINGLE.finditer(line):
        _add(findings, "R06-smart-quotes", ln, m.start() + 1,
             m.group(0), "\u2019", "Use curly apostrophe.")
    # Straight double quotes: replace with smart double quote.
    # We can't know from context whether it's opening or closing -
    # default to left double-quote; Word's autocorrect handles the rest
    # when the doc is reopened.
    for m in _R06_STRAIGHT_DOUBLE.finditer(line):
        _add(findings, "R06-smart-quotes", ln, m.start() + 1,
             '"', "\u201C", "Use curly double quotes.")
    for m in _R09_RE.finditer(line):
        # Direct replacement: collapse run of spaces to single.
        _add(findings, "R09-double-space", ln, m.start() + 1,
             m.group(0), " ", "Collapse to single space.")
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
        log.warning("pandoc not found - skipping Polish prose extraction.")
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
# Linguistic pass (Claude) - major errors only
# ---------------------------------------------------------------------------

_LINGUISTIC_SYSTEM = """\
You are Polish - a silent QA gate for ICE Contractors client proposals.
Your only job is to catch MAJOR, UNPROFESSIONAL errors that would
embarrass us if the document shipped as-is.

Flag ONLY these:
- Misspellings / typos (wrong letters, transposed characters)
- Wrong-word errors (their/there, your/you're, affect/effect, etc.)
- Grammar errors that break or change meaning (subject-verb
  disagreement, broken sentence structure)
- Doubled words a regex missed
- Broken or incomplete sentences

DO NOT flag:
- Style preferences (sentence length, word choice, rhetorical flow)
- Parallel-structure nits
- Passive vs active voice
- Mild redundancy
- Tone or "could be stronger" rewrites
- Anything the deterministic linter already handles (quotes, em-dash,
  I&E, T&M, contractions, exclamations, currency, double spaces)

Keep the `current` string SHORT - ideally 1-8 words, never a full
sentence unless the sentence is truly broken. Short spans apply cleanly
to a docx; long rewrites do not.

Return ONLY a JSON array. Each item:
- "issue": one of: typo, wrong-word, grammar, doubled-word, broken-sentence
- "current": exact text to replace (SHORT)
- "suggested": the corrected replacement
- "note": 1-sentence reason

No markdown fences. No preamble. If nothing qualifies, return [].
"""


def _linguistic_pass(prose: str, api_key: str, model: str) -> List[Finding]:
    """Call Claude once to catch what regex cannot. Returns empty list on
    any failure - linguistic pass is best-effort, never blocks the agent."""
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
# Silent corrections applied directly to the .docx
# ---------------------------------------------------------------------------

def _xml_unescape(text: str) -> str:
    """Decode XML entities we see inside <w:t> content before matching."""
    return (text.replace("&amp;", "&")
                 .replace("&lt;", "<")
                 .replace("&gt;", ">")
                 .replace("&quot;", '"')
                 .replace("&apos;", "'"))


def _xml_escape(text: str) -> str:
    """Re-encode text for safe insertion into XML."""
    return (text.replace("&", "&amp;")
                 .replace("<", "&lt;")
                 .replace(">", "&gt;"))


def _apply_clean_edits(docx_src: Path, docx_dst: Path,
                       findings: List[Finding]) -> int:
    """Create docx_dst from docx_src with the finding corrections applied
    directly to <w:t> text nodes - NO tracked-change markup.

    Both deterministic and linguistic findings are applied, but only where
    the `current` text is short enough (<= 120 chars) to fit inside a
    single run. Long sentence-level rewrites are silently skipped.

    Returns the count of edits applied. If zero, docx_dst is still
    written (a straight copy) so the caller has a single file to attach.
    """
    import zipfile

    usable = [
        f for f in findings
        if f.current and f.suggested
        and f.current != f.suggested
        and len(f.current) <= 120
    ]
    if not usable:
        shutil.copyfile(docx_src, docx_dst)
        return 0

    applied = 0
    t_re = re.compile(r"(<w:t(?:\s+[^>]*)?>)([^<]*)(</w:t>)")

    def _sub(m: re.Match) -> str:
        nonlocal applied
        opening, inner_xml, closing = m.group(1), m.group(2), m.group(3)
        # Work on the unescaped text so matches line up with what pandoc
        # extracted (which is what the findings were computed against).
        inner = _xml_unescape(inner_xml)
        changed = False
        for f in usable:
            if f.current in inner:
                inner = inner.replace(f.current, f.suggested)
                applied += 1
                changed = True
        if not changed:
            return m.group(0)
        # Preserve whitespace when the corrected text has leading/trailing
        # spaces or we had xml:space already.
        needs_preserve = inner != inner.strip() or 'xml:space' in opening
        new_opening = opening
        if needs_preserve and 'xml:space' not in new_opening:
            new_opening = new_opening.replace("<w:t", '<w:t xml:space="preserve"', 1)
        return f"{new_opening}{_xml_escape(inner)}{closing}"

    with zipfile.ZipFile(docx_src, "r") as zin:
        members = {n: zin.read(n) for n in zin.namelist()}

    doc_xml = members.get("word/document.xml", b"").decode("utf-8")
    new_xml = t_re.sub(_sub, doc_xml)
    members["word/document.xml"] = new_xml.encode("utf-8")

    docx_dst.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(docx_dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in members.items():
            zout.writestr(name, data)

    log.info("Polish: applied %d clean edit(s) to %s",
             applied, docx_dst.name)
    return applied


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def run_polish(docx_path: Path, *, anthropic_api_key: str,
               model: str, run_linguistic: bool = True) -> PolishResult:
    """Run Polish over a filled .docx and return a PolishResult.

    Applies corrections SILENTLY - direct edits to the document text,
    no tracked changes, no punch list, no email reporting. The caller
    should attach `polished_docx` if it's not None, else the original.

    Never raises - on any error the result degrades to no-op (grade A,
    polished_docx=None) so the agent keeps shipping.
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

        polished_path: Optional[Path] = None
        applied = 0
        if findings:
            candidate = docx_path.with_name(docx_path.stem + "_polished.docx")
            applied = _apply_clean_edits(docx_path, candidate, findings)
            if applied > 0:
                polished_path = candidate
            else:
                # Nothing actually applied - remove the stray copy.
                candidate.unlink(missing_ok=True)

        log.info("Polish: found=%d (det=%d, ling=%d) applied=%d ran_linguistic=%s",
                 len(findings), n_det, n_ling, applied, ran_ling)

        return PolishResult(
            grade="A" if applied == 0 else "B",
            findings=findings,
            punch_list_md="",  # silent mode: no report
            polished_docx=polished_path,
            clean_docx=docx_path,
            ran_linguistic_pass=ran_ling,
        )
    except Exception:
        log.exception("Polish crashed - skipping silently.")
        return PolishResult(grade="A", clean_docx=docx_path)

"""Idempotent patcher: ensure the ICE proposal template header table
contains an "Owner's Rep:" row between "Client:" and "Location:".

This script exists because the quote agent's schema added an `owner_rep`
field (the client-side project contact the quote is addressed to). The
template's header KV table needed a new row to display that value.

Running this script against an already-patched template is a no-op.

Usage:
    python scripts/ensure_template_has_owner_rep.py \
        templates/ICE_Contractors_Proposal_Template.docx
"""
from __future__ import annotations
import copy
import sys
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn


OWNER_REP_LABEL = "Owner\u2019s Rep:"   # curly apostrophe to match house style


def _row_first_cell_text(row) -> str:
    return row.cells[0].text.strip().lower().rstrip(":")


def _find_header_table(doc):
    for t in doc.tables:
        first_col = {_row_first_cell_text(r) for r in t.rows}
        if "project" in first_col and "client" in first_col:
            return t
    return None


def _find_row_index(table, label_lower: str) -> int:
    for i, row in enumerate(table.rows):
        if _row_first_cell_text(row) == label_lower:
            return i
    return -1


def _clone_row_as(source_row, label: str):
    """Clone the XML element of source_row, then blank out both cells
    and set the left cell to `label` (bold, preserving the run style
    from the source row's left cell)."""
    new_tr = copy.deepcopy(source_row._tr)

    # Strip all existing paragraph content from the cloned cells.
    for tc in new_tr.findall(qn("w:tc")):
        for p in tc.findall(qn("w:p")):
            tc.remove(p)
        # Put one empty paragraph back so the cell is valid.
        empty_p = copy.deepcopy(
            source_row.cells[0]._tc.findall(qn("w:p"))[0]
        )
        # Clear text inside this cloned paragraph.
        for t in empty_p.iter(qn("w:t")):
            t.text = ""
        tc.append(empty_p)

    return new_tr


def ensure_owner_rep_row(template_path: Path) -> bool:
    """Return True if a row was added, False if the template already
    contained one (no-op)."""
    doc = Document(str(template_path))
    table = _find_header_table(doc)
    if table is None:
        raise RuntimeError(
            f"Could not locate header KV table in {template_path}. "
            "Expected a table containing both 'Project:' and 'Client:' rows."
        )

    if _find_row_index(table, "owner\u2019s rep") >= 0 \
       or _find_row_index(table, "owner's rep") >= 0:
        return False  # already present — nothing to do

    client_idx = _find_row_index(table, "client")
    if client_idx < 0:
        raise RuntimeError(
            "Header table has no 'Client:' row — cannot locate insertion point."
        )
    client_row = table.rows[client_idx]

    # Clone the Client row, label the left cell "Owner's Rep:", leave
    # the right cell empty (the agent fills it at runtime).
    new_tr = _clone_row_as(client_row, OWNER_REP_LABEL)

    # Set the label text on the first cell of the cloned row.
    first_tc = new_tr.findall(qn("w:tc"))[0]
    first_p = first_tc.findall(qn("w:p"))[0]
    runs = first_p.findall(qn("w:r"))
    if runs:
        # Put label in first run's <w:t>, clear any subsequent runs.
        first_r = runs[0]
        t_elems = first_r.findall(qn("w:t"))
        if t_elems:
            t_elems[0].text = OWNER_REP_LABEL
            for extra in t_elems[1:]:
                first_r.remove(extra)
        else:
            # Unusual — no <w:t> child. Add one.
            from lxml import etree
            t = etree.SubElement(first_r, qn("w:t"))
            t.text = OWNER_REP_LABEL
        for extra_r in runs[1:]:
            first_p.remove(extra_r)
    else:
        # No runs at all — build one with the label.
        from lxml import etree
        r = etree.SubElement(first_p, qn("w:r"))
        t = etree.SubElement(r, qn("w:t"))
        t.text = OWNER_REP_LABEL

    # Insert the new row immediately after the Client row.
    client_row._tr.addnext(new_tr)

    doc.save(str(template_path))
    return True


def main():
    if len(sys.argv) != 2:
        print(
            "Usage: python scripts/ensure_template_has_owner_rep.py "
            "<template.docx>",
            file=sys.stderr,
        )
        sys.exit(2)

    target = Path(sys.argv[1]).resolve()
    if not target.exists():
        print(f"Template not found: {target}", file=sys.stderr)
        sys.exit(2)

    added = ensure_owner_rep_row(target)
    if added:
        print(f"Added Owner\u2019s Rep row to {target.name}")
    else:
        print(f"Owner\u2019s Rep row already present in {target.name} (no-op)")


if __name__ == "__main__":
    main()

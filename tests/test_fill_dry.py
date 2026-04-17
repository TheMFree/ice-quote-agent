"""Offline dry-run: exercises the template filler with a fabricated
QuoteData object. Does not call Claude or touch any mailbox.

Run:  python tests/test_fill_dry.py
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from agent.filler import fill_template
from agent.schema import QuoteData, ScopeCategory


def main():
    data = QuoteData(
        document_type="QUOTATION",
        project="Baron #3 Installation — Warehouse Electrical Upgrade",
        client="Hitachi Energy",
        location="Richland, MS",
        proposal_date="April 17, 2026",
        quote_no="Q-2026-0417",
        scope_intro=(
            "ICE Contractors, Inc. proposes to furnish all labor, equipment, and "
            "supervision required to complete the electrical installation for the "
            "Baron #3 Warehouse Electrical Upgrade at the Hitachi Energy facility "
            "in Richland, MS."
        ),
        scope_bullets=[
            "2000A QED-2 switchboard (owner-furnished) — set in place and connect.",
            "1200A MDP fed from switchboard via cable tray.",
            "600A and 400A sub-panels with associated feeders.",
            "75 kVA dry-type transformer and 100A sub-panel.",
            "HVAC disconnects (3 x 200A) and unit feeders.",
            "Megger testing, phasing, and commissioning.",
        ],
        material_amount="$269,270.00",
        labor_equipment_amount="$202,363.00",
        total_amount="$471,633.00",
        estimated_duration="Approximately 6 weeks",
        work_schedule="4-10s (Monday–Thursday)",
        extraction_notes="Dry-run test.",
    )

    template = ROOT / "templates" / "ICE_Contractors_Proposal_Template.docx"
    out = ROOT / "output" / "dry_run_test.docx"
    out.parent.mkdir(parents=True, exist_ok=True)
    fill_template(template, data, out)
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()

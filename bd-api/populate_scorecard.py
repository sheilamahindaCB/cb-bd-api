"""
populate_scorecard.py
Reads ZREP_Zimbabwe_GoNoGo.xlsx, writes assessed data into the correct cells,
and saves a new file named <short_name>_GoNoGo.xlsx.

Input: JSON with the structure output by the CB-PNI-BD-Assessment skill.
Usage: python populate_scorecard.py assessment.json PakistanDISCOIV
"""

import sys
import json
import shutil
import os
from datetime import datetime
from openpyxl import load_workbook

TEMPLATE = os.path.join(os.path.dirname(__file__), "templates", "ZREP_Zimbabwe_GoNoGo.xlsx")


def populate(data: dict, short_name: str, output_dir: str = "outputs") -> str:
    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, f"{short_name}_GoNoGo.xlsx")
    shutil.copy2(TEMPLATE, out_path)

    wb = load_workbook(out_path)
    ws = wb["Go-NoGo Scorecard"]

    def safe_set(cell_ref, value):
        """Only write to non-merged cells."""
        from openpyxl.cell.cell import MergedCell
        cell = ws[cell_ref]
        if not isinstance(cell, MergedCell):
            cell.value = value

    bid = data.get("bid_identification", {})
    scoring = data.get("strategic_fit", [])
    experts = data.get("key_expert_coverage", [])
    partners = data.get("partner_shortlist", [])
    consortium = data.get("recommended_consortium", [])

    # ── Section A — Bid Identification (column C, rows 3-17) ──────────────────
    field_map = {
        "opportunity_name":         "C5",
        "client":                   "C6",
        "funder":                   "C7",
        "funder_type":              "C8",
        "country":                  "C9",
        "sector":                   "C10",
        "procurement_method":       "C11",
        "reference_no":             "C12",
        "submission_deadline":      "C13",
        "estimated_contract_value": "C16",
        "cb_share_pct":             "C18",
    }
    for key, cell in field_map.items():
        val = bid.get(key)
        if val is not None:
            if key == "submission_deadline" and isinstance(val, str):
                try:
                    val = datetime.strptime(val, "%Y-%m-%d")
                except ValueError:
                    pass
            safe_set(cell, val)

    # ── Section B — Strategic Fit Scoring (rows 22-31, col F = score) ─────────
    # scoring list: [{criterion, score, weight, justification}, ...]
    b_start = 22
    for i, item in enumerate(scoring[:10]):
        row = b_start + i
        score = item.get("score")
        notes = item.get("justification", "")
        if score is not None:
            safe_set(f"F{row}", score)
        if notes:
            safe_set(f"I{row}", notes)

    # ── Section C — Key Expert Coverage (rows 36-45, col E = CB score) ────────
    c_start = 36
    for i, item in enumerate(experts[:10]):
        row = c_start + i
        cb_score = item.get("cb_score")
        filled_by = item.get("filled_by", "")
        comment = item.get("comment", "")
        if cb_score is not None:
            safe_set(f"E{row}", cb_score)
        if filled_by:
            safe_set(f"G{row}", filled_by)
        if comment:
            safe_set(f"H{row}", comment)

    # ── Section D — Partner Shortlist (rows 50-59) ────────────────────────────
    d_start = 50
    for i, item in enumerate(partners[:10]):
        row = d_start + i
        for col, key in [("B", "partner_name"), ("C", "role"), ("D", "past_jv"),
                         ("F", "sector_score"), ("G", "region_score"),
                         ("H", "local_presence_score"), ("I", "capacity_score"),
                         ("L", "contact"), ("M", "rationale")]:
            val = item.get(key)
            if val is not None:
                safe_set(f"{col}{row}", val)

    # ── Section E — Recommended Consortium (rows 63-67) ──────────────────────
    e_start = 63
    for i, item in enumerate(consortium[:5]):
        row = e_start + i
        role = item.get("role", "")
        firm = item.get("firm", "")
        scope = item.get("scope", "")
        rationale = item.get("rationale", "")
        if role:
            safe_set(f"B{row}", role)
        if firm:
            safe_set(f"C{row}", firm)
        # Column D is merged — write scope into C as "Firm — Scope"
        if scope and firm:
            safe_set(f"C{row}", f"{firm} — {scope}")
        if rationale:
            safe_set(f"I{row}", rationale)

    wb.save(out_path)
    print(f"Scorecard saved: {out_path}")
    return out_path


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python populate_scorecard.py assessment.json <short_name>")
        sys.exit(1)
    with open(sys.argv[1]) as f:
        data = json.load(f)
    populate(data, sys.argv[2])

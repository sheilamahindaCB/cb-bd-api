"""
populate_memo.py
Reads ZREP_Zimbabwe_Approval_Memo.docx, replaces all placeholder fields
with assessed data, and saves a new file named <short_name>_ApprovalMemo.docx.

Input: JSON with the structure output by the CB-PNI-BD-Assessment skill.
Usage: python populate_memo.py assessment.json PakistanDISCOIV
"""

import sys
import json
import shutil
import os
from datetime import datetime
from docx import Document
from docx.oxml.ns import qn

TEMPLATE = os.path.join(os.path.dirname(__file__), "templates", "ZREP_Zimbabwe_Approval_Memo.docx")


def replace_text_in_paragraph(para, replacements: dict):
    """Replace placeholder text in a paragraph, preserving formatting of the first run."""
    full_text = "".join(run.text for run in para.runs)
    new_text = full_text
    for placeholder, value in replacements.items():
        new_text = new_text.replace(placeholder, str(value) if value else "")
    if new_text != full_text:
        if para.runs:
            para.runs[0].text = new_text
            for run in para.runs[1:]:
                run.text = ""


def replace_in_table(table, replacements: dict):
    for row in table.rows:
        for cell in row.cells:
            for para in cell.paragraphs:
                replace_text_in_paragraph(para, replacements)
            for nested_table in cell.tables:
                replace_in_table(nested_table, replacements)


def populate(data: dict, short_name: str, output_dir: str = "outputs") -> str:
    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, f"{short_name}_ApprovalMemo.docx")
    shutil.copy2(TEMPLATE, out_path)

    doc = Document(out_path)

    bid = data.get("bid_identification", {})
    scoring = data.get("strategic_fit", [])
    experts = data.get("key_expert_coverage", [])
    consortium = data.get("recommended_consortium", [])
    decision_log = data.get("decision_log", {})
    gaps = data.get("top_gaps", [])
    next_steps = data.get("next_steps", [])

    today = datetime.today().strftime("%-d %B %Y")
    score = decision_log.get("score", "")
    decision = decision_log.get("decision", "")

    # Build gap text
    gap_bullets = ""
    for i, gap in enumerate(gaps[:2], 1):
        gap_bullets += f"\n{i}. {gap.get('gap', '')} — {gap.get('mitigation', '')}"

    # Build consortium rows text
    consortium_text = ""
    for item in consortium:
        consortium_text += f"\n{item.get('role','')}: {item.get('firm','')} — {item.get('scope','')}"

    # Build next steps text
    steps_text = ""
    for step in next_steps[:5]:
        steps_text += f"\n- {step}"

    replacements = {
        # Header fields
        "21 May 2026":                          today,
        "Go / No-Go Recommendation — ZREP-TA Solar PV Procurement TA, Zimbabwe":
            f"Go / No-Go Recommendation — {bid.get('opportunity_name', '')}",
        "David Ndungu, Associate, CrossBoundary Power & Infrastructure":
            "BD Team, CrossBoundary Power & Infrastructure",
        "Transaction Advisory Services for the Procurement of 370 MW Solar PV\nProjects under the ZREP-TA Project":
            bid.get("opportunity_name", ""),

        # Section A table values
        "Zimbabwe Electricity Transmission\n                                      and Distribution Company (ZETDC)":
            bid.get("client", ""),
        "World Bank — IDA / SRMI ESMAP grant\n                                      TF0D0261; Project P511160":
            bid.get("funder", ""),
        "Zimbabwe / Southern Africa / SADC":    bid.get("country", ""),
        "Solar IPP / Auction TA — 370 MW\n                                      Solar PV (140 MW Solar Park + 230\n                                      MW Substation-Based)":
            bid.get("sector", ""),
        "QCBS":                                 bid.get("procurement_method", "QCBS"),
        "ZREP-TA/TA/01/2025 (Borrower bid\n                                      ref: ZW-ZETDC-533247-CS-QCBS)":
            bid.get("reference_no", ""),
        "May 29, 2026 16:30 CAT":               str(bid.get("submission_deadline", "")),
        "US$2,500,000 (CB share ~40%)":         f"US${bid.get('estimated_contract_value', '')} (CB share ~{int(bid.get('cb_share_pct', 0.4)*100)}%)",
        "World Bank STD-EOI (Firms\n                                      Selection)":
            bid.get("procurement_method", ""),

        # Section B decision
        "Decision: GO (Consortium-Mandatory — lock partners before kickoff)":
            f"Decision: {decision}",
        "3.92/5 (PASS, threshold ≥3.5)":        f"{score}/5",
        "43.3%":                                 f"{data.get('cb_coverage_pct', '')}%",
    }

    # Apply replacements to all paragraphs
    for para in doc.paragraphs:
        replace_text_in_paragraph(para, replacements)

    # Apply replacements to all tables
    for table in doc.tables:
        replace_in_table(table, replacements)

    doc.save(out_path)
    print(f"Approval memo saved: {out_path}")
    return out_path


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python populate_memo.py assessment.json <short_name>")
        sys.exit(1)
    with open(sys.argv[1]) as f:
        data = json.load(f)
    populate(data, sys.argv[2])

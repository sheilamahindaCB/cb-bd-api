"""
populate_eoi.py
Reads ZREP_Zimbabwe_EOI_Draft.docx, replaces all placeholder fields
with assessed data, and saves a new file named <short_name>_EOI_Draft.docx.

Input: JSON with the structure output by the CB-PNI-BD-Assessment skill.
Usage: python populate_eoi.py assessment.json PakistanDISCOIV
"""

import sys
import json
import shutil
import os
from datetime import datetime
from docx import Document

TEMPLATE = os.path.join(os.path.dirname(__file__), "templates", "ZREP_Zimbabwe_EOI_Draft.docx")


def replace_text_in_paragraph(para, replacements: dict):
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
    out_path = os.path.join(output_dir, f"{short_name}_EOI_Draft.docx")
    shutil.copy2(TEMPLATE, out_path)

    doc = Document(out_path)

    bid = data.get("bid_identification", {})
    consortium = data.get("recommended_consortium", [])
    decision_log = data.get("decision_log", {})

    today = datetime.today().strftime("%-d %B %Y")

    # Build consortium string: "CB LLC, K&M Advisors, ..."
    consortium_firms = ", ".join(
        item.get("firm", "") for item in consortium if item.get("firm")
    )

    replacements = {
        "21 May 2026":                  today,
        "Transaction Advisory Services for the Procurement of 370 MW Solar PV Projects under the ZREP-TA Project":
            bid.get("opportunity_name", ""),
        "CrossBoundary LLC, K&M Advisors,\nMulticonsult UK Ltd, and Hunton Andrews Kurth LLP":
            consortium_firms,
        "CrossBoundary LLC, in consortium with K&M Advisors,\nMulticonsult UK Ltd, and Hunton Andrews Kurth LLP":
            f"CrossBoundary LLC, in consortium with {consortium_firms}",
        "ZREP-TA/TA/01/2025":           bid.get("reference_no", ""),
        "ZW-ZETDC-533247-CS-QCBS":      bid.get("reference_no", ""),
        "P511160":                       bid.get("funder_project_id", ""),
        "ZETDC's Request for Expressions of Interest (Ref: ZREP-TA/TA/01/2025)":
            f"the Request for Expressions of Interest (Ref: {bid.get('reference_no', '')})",
        "370 MW of Solar PV Projects under the ZREP-TA Project":
            bid.get("opportunity_name", ""),
        "Zimbabwe Electricity Transmission and Distribution Company (ZETDC)":
            bid.get("client", ""),
        "the CrossBoundary-led consortium (CrossBoundary LLC, K&M Advisors LLC,\nMulticonsult UK Ltd, and Hunton Andrews Kurth LLP)":
            f"the CrossBoundary-led consortium ({consortium_firms})",
    }

    for para in doc.paragraphs:
        replace_text_in_paragraph(para, replacements)

    for table in doc.tables:
        replace_in_table(table, replacements)

    doc.save(out_path)
    print(f"EOI draft saved: {out_path}")
    return out_path


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python populate_eoi.py assessment.json <short_name>")
        sys.exit(1)
    with open(sys.argv[1]) as f:
        data = json.load(f)
    populate(data, sys.argv[2])

"""
api.py
Flask API that accepts a JSON assessment payload and returns
populated Excel and Word files as downloads.

Endpoints:
  POST /generate/scorecard   → returns <short_name>_GoNoGo.xlsx
  POST /generate/memo        → returns <short_name>_ApprovalMemo.docx
  POST /generate/eoi         → returns <short_name>_EOI_Draft.docx
  POST /generate/all         → returns a ZIP with all three files

Payload (JSON body):
{
  "short_name": "PakistanDISCOIV",
  "bid_identification": { ... },
  "strategic_fit": [ ... ],
  "key_expert_coverage": [ ... ],
  "partner_shortlist": [ ... ],
  "recommended_consortium": [ ... ],
  "decision_log": { "score": 3.32, "decision": "CONDITIONAL GO" },
  "top_gaps": [ ... ],
  "next_steps": [ ... ],
  "cb_coverage_pct": 43
}

Run locally:
  pip install flask python-docx openpyxl
  python api.py

Deploy to Azure Functions / AWS Lambda / Railway / Render:
  See README.md for deployment steps.
"""

import os
import io
import zipfile
from flask import Flask, request, jsonify, send_file

from populate_scorecard import populate as populate_scorecard
from populate_memo import populate as populate_memo
from populate_eoi import populate as populate_eoi

app = Flask(__name__)
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "outputs")


def get_payload():
    data = request.get_json(force=True, silent=True)
    if not data:
        return None, jsonify({"error": "Invalid or missing JSON body"}), 400
    short_name = data.get("short_name", "output").replace(" ", "_")
    return data, short_name, None


@app.route("/generate/scorecard", methods=["POST"])
def gen_scorecard():
    data, short_name, err = get_payload()
    if err:
        return err
    path = populate_scorecard(data, short_name, OUTPUT_DIR)
    return send_file(path, as_attachment=True,
                     download_name=os.path.basename(path),
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/generate/memo", methods=["POST"])
def gen_memo():
    data, short_name, err = get_payload()
    if err:
        return err
    path = populate_memo(data, short_name, OUTPUT_DIR)
    return send_file(path, as_attachment=True,
                     download_name=os.path.basename(path),
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/generate/eoi", methods=["POST"])
def gen_eoi():
    data, short_name, err = get_payload()
    if err:
        return err
    path = populate_eoi(data, short_name, OUTPUT_DIR)
    return send_file(path, as_attachment=True,
                     download_name=os.path.basename(path),
                     mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.route("/generate/all", methods=["POST"])
def gen_all():
    data, short_name, err = get_payload()
    if err:
        return err
    paths = [
        populate_scorecard(data, short_name, OUTPUT_DIR),
        populate_memo(data, short_name, OUTPUT_DIR),
        populate_eoi(data, short_name, OUTPUT_DIR),
    ]
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in paths:
            zf.write(path, os.path.basename(path))
    zip_buffer.seek(0)
    return send_file(zip_buffer, as_attachment=True,
                     download_name=f"{short_name}_BD_Outputs.zip",
                     mimetype="application/zip")


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)

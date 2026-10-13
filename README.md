# CB P&I BD — Template Population API

Takes a JSON assessment payload from the CB-PNI-BD-Assessment Claude skill
and returns populated Excel and Word files as downloads.

---

## Files

| File | Purpose |
|---|---|
| `api.py` | Flask API server — run this |
| `populate_scorecard.py` | Fills ZREP_Zimbabwe_GoNoGo.xlsx |
| `populate_memo.py` | Fills ZREP_Zimbabwe_Approval_Memo.docx |
| `populate_eoi.py` | Fills ZREP_Zimbabwe_EOI_Draft.docx |
| `templates/` | Master ZREP template files — never modify these |
| `outputs/` | Generated files land here — one subfolder per bid |
| `sample_assessment.json` | Example JSON payload for testing |

---

## Local setup (one-time)

```bash
pip install -r requirements.txt
python api.py
```

Server runs at http://localhost:5000

---

## API endpoints

### Generate all three files (ZIP download)
```
POST /generate/all
Content-Type: application/json
Body: <assessment JSON>
```

### Generate individual files
```
POST /generate/scorecard   → PakistanDISCOIV_GoNoGo.xlsx
POST /generate/memo        → PakistanDISCOIV_ApprovalMemo.docx
POST /generate/eoi         → PakistanDISCOIV_EOI_Draft.docx
```

### Health check
```
GET /health
```

---

## Test locally

```bash
curl -X POST http://localhost:5000/generate/all \
  -H "Content-Type: application/json" \
  -d @sample_assessment.json \
  --output test_outputs.zip
```

---

## JSON payload structure

Claude outputs this JSON at the end of every assessment.
Copy it and POST it to /generate/all.

```json
{
  "short_name": "PakistanDISCOIV",
  "bid_identification": {
    "opportunity_name": "...",
    "client": "...",
    "funder": "...",
    "funder_type": "MDB",
    "country": "...",
    "sector": "...",
    "procurement_method": "QCBS",
    "reference_no": "...",
    "submission_deadline": "2026-12-15",
    "estimated_contract_value": 3500000,
    "cb_share_pct": 0.35
  },
  "strategic_fit": [
    {"criterion": "...", "score": 4, "weight": 0.15, "justification": "..."}
  ],
  "key_expert_coverage": [
    {"role": "...", "cb_score": 3, "gap": "None", "filled_by": "...", "comment": "..."}
  ],
  "partner_shortlist": [
    {"partner_name": "...", "role": "...", "past_jv": "Yes/No",
     "sector_score": 5, "region_score": 3, "local_presence_score": 2,
     "capacity_score": 5, "contact": "...", "rationale": "..."}
  ],
  "recommended_consortium": [
    {"role": "...", "firm": "...", "scope": "...", "rationale": "..."}
  ],
  "decision_log": {
    "score": 3.32,
    "decision": "CONDITIONAL GO"
  },
  "top_gaps": [
    {"gap": "...", "mitigation": "..."}
  ],
  "next_steps": ["...", "..."],
  "cb_coverage_pct": 43
}
```

---

## Deploy to the cloud (non-technical team option)

The easiest free option is **Railway** (railway.app):

1. Push this folder to a GitHub repository
2. Go to railway.app, create a new project, connect the GitHub repo
3. Railway auto-detects Python and deploys it
4. Copy the deployment URL (e.g. https://cb-bd-api.railway.app)
5. Update the SKILL.md to POST to that URL after every assessment

Team members then never need to touch Python — Claude runs the assessment,
POSTs the JSON to the API, and gets download links back automatically.

---

## Update templates

When the ZREP master templates change, replace the files in `templates/`
and redeploy. The population scripts will pick up the new layout automatically
as long as the cell structure stays the same.

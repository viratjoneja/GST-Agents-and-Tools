# GST Agents and Tools

**GSTR-1 Agentic Reconciliation Pipeline**
Built by Y K Joneja & Co., Faridabad | Part of TrueGST.ai

## Architecture

4-agent sequential CrewAI pipeline:

1. **Reconciliation Analyst** — Runs the GSTR-1 vs EWB engine
2. **Investigation Specialist** — Classifies discrepancies (threshold exemptions, export invoices, job work)
3. **Compliance Officer** — Applies GST law, calculates revised filing position
4. **Report Writer** — Produces filing readiness report

## Features

- 3-way reconciliation: Sales Register × E-Way Bill × E-Invoice
- Handles separate file inputs (Sales + EWB + E-Invoice)
- Non-standard header detection (title row skip)
- Delivery Challan separation (job work vs taxable supply)
- EWB threshold rules: inter-state ₹50,000, Haryana intra-state ₹50,000
- Export invoice detection
- Outputs: Excel workbook + Word filing readiness report

## Files

| File | Purpose |
|------|---------|
| `recon_engine_api.py` | Reconciliation engine — loads sales + EWB + e-invoice, matches, outputs Excel |
| `intelligence.py` | Claude-based config extraction from file structure |
| `gstr1_crew_agent.py` | CrewAI agents, tasks, tools, and crew orchestration |
| `requirements.txt` | Dependencies |

## Usage

See `gstr1_crew_agent.py` — designed to run in Google Colab.
Set `GOOGLE_API_KEY` in Colab Secrets before running.

## Clients Tested

- Gripple Hanger Joiner Systems (I) Pvt. Ltd. — January 2026
- Promptech Industrial Products Pvt. Ltd. — March 2026

## Tech Stack

- Python 3.12
- CrewAI 0.80+
- Gemini 2.5 Flash
- pandas, openpyxl, python-docx

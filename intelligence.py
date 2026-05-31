
import os, json, httpx

ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages"
MODEL             = "claude-haiku-4-5-20251001"
MAX_TOKENS        = 1000

def _headers():
    return {
        "Content-Type":      "application/json",
        "x-api-key":         os.environ.get("ANTHROPIC_API_KEY", ""),
        "anthropic-version": "2023-06-01",
        "anthropic-beta":    "prompt-caching-2024-07-31",
    }

_SYS_CONFIG = """You are a GST data analyst for an Indian CA/Advocate practice.
Analyse the Excel file structure and identify which sheets contain the Sales Register
and the E-Way Bill Register, then extract client metadata.

Rules:
- Sales Register: columns like Invoice Number, Invoice Date, Receiver Name,
  GSTIN/UIN of Recipient, Taxable Value, Invoice Value, Rate, Place of Supply.
- EWB sheet: columns like From GSTIN, From GSTIN Info, Doc.No, Doc.Date,
  Other Party GSTIN, Assessable Value, Total Invoice Value, EWB No.
- CLIENT_NAME: from "From GSTIN Info" column of EWB sheet, text before double spaces.
- GSTIN: from "From GSTIN" column of EWB sheet, first non-null value.
- TAX_PERIOD: infer from Invoice Date column — most common month/year as "Month YYYY".
- If a value cannot be determined confidently, return empty string.

Respond ONLY with valid JSON. No preamble, no markdown fences."""

async def extract_config(probe: dict) -> dict:
    user_message = f"""Analyse this Excel file structure and extract the configuration:

Sheet names: {probe["sheet_names"]}

Sheet details:
{json.dumps(probe["sheets"], indent=2, ensure_ascii=False)}

Return JSON with exactly these keys:
{{
  "client_name": "",
  "gstin": "",
  "tax_period": "",
  "sheet_sales": "",
  "sheet_ewb": ""
}}"""

    payload = {
        "model": MODEL, "max_tokens": MAX_TOKENS,
        "system": [{"type":"text","text":_SYS_CONFIG,"cache_control":{"type":"ephemeral"}}],
        "messages": [{"role":"user","content":user_message}]
    }
    async with httpx.AsyncClient(timeout=30) as client:
        resp = await client.post(ANTHROPIC_API_URL, headers=_headers(), json=payload)
        resp.raise_for_status()
        raw = resp.json()["content"][0]["text"].strip()
    if raw.startswith("```"):
        raw = raw.split("```")[1]
        if raw.startswith("json"): raw = raw[4:]
    return json.loads(raw.strip())

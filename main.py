"""
GST Agents API — FastAPI entry point
"""
import os, base64, json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import asyncio

from crew_runner import run_gst_analysis

app = FastAPI(title="GST Agents API",
              description="GSTR-1 Agentic Reconciliation — Y K Joneja & Co.",
              version="1.0.0")

app.add_middleware(CORSMiddleware,
    allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

@app.get("/health")
def health():
    return {"status": "ok", "service": "gst-agents-tools", "version": "1.0.0"}

@app.post("/reconcile")
async def reconcile(
    sales_file:        UploadFile = File(...),
    ewb_file:          UploadFile = File(...),
    einv_file:         UploadFile = File(None),
    client_name:       str = Form(...),
    gstin:             str = Form(...),
    tax_period:        str = Form(...),
    sheet_sales:       str = Form(...),
    sheet_ewb:         str = Form(...),
    einv_applicable:   str = Form("no"),
    intra_threshold:   str = Form("50000"),
    sales_header_row:  str = Form("1"),
):
    sales_bytes = await sales_file.read()
    ewb_bytes   = await ewb_file.read()
    einv_bytes  = await einv_file.read() if einv_file else None

    config = {
        "client_name":           client_name.strip(),
        "gstin":                 gstin.strip().upper(),
        "tax_period":            tax_period.strip(),
        "sheet_sales":           sheet_sales.strip(),
        "sheet_ewb":             sheet_ewb.strip(),
        "cdn_start_row":         None,
        "sales_header_row":      int(sales_header_row) if sales_header_row.isdigit() else 1,
        "einv_applicable":       einv_applicable.lower() in ("yes","true","1"),
        "intra_state_threshold": int(intra_threshold) if intra_threshold.isdigit() else 50000,
    }

    try:
        result = await asyncio.to_thread(
            run_gst_analysis, sales_bytes, ewb_bytes, config, einv_bytes)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    return JSONResponse(content={
        "stats":       result["stats"],
        "report_text": result["report_text"],
        "excel_b64":   result["excel_b64"],
        "docx_b64":    result["docx_b64"],
    })

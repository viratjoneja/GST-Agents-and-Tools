"""
crew_runner.py — GSTR-1 Agentic Reconciliation
Stateless per-request execution. Safe for concurrent web requests.
"""

import os, json, base64, asyncio, contextvars, concurrent.futures
import pandas as pd
from io import BytesIO, StringIO
from crewai import Agent, Task, Crew, Process, LLM
from crewai.tools import BaseTool
from langsmith import traceable
from recon_engine_api import run_reconciliation


def _setup_tracing():
    try:
        from langsmith.integrations.otel import OtelSpanProcessor
        from opentelemetry import trace
        from opentelemetry.sdk.trace import TracerProvider
        from opentelemetry.instrumentation.crewai import CrewAIInstrumentor
        current = trace.get_tracer_provider()
        tp = current if isinstance(current, TracerProvider) else TracerProvider()
        if not isinstance(current, TracerProvider):
            trace.set_tracer_provider(tp)
        tp.add_span_processor(OtelSpanProcessor())
        CrewAIInstrumentor().instrument(tracer_provider=tp)
        print("✔ LangSmith tracing configured")
    except Exception as ex:
        print(f"⚠ LangSmith tracing not configured: {ex}")

_setup_tracing()


class RunState:
    def __init__(self, sales_bytes, ewb_bytes, config, einv_bytes=None):
        self.file_bytes  = sales_bytes
        self.ewb_bytes   = ewb_bytes
        self.einv_bytes  = einv_bytes
        self.config      = config
        self.excel_bytes = None
        self.stats       = None


class RunReconciliationTool(BaseTool):
    name: str        = "Run GST Reconciliation"
    description: str = "Run the GSTR-1 vs EWB reconciliation engine. Always call FIRST."
    state: object    = None
    class Config: arbitrary_types_allowed = True
    def _run(self, instruction: str = "run") -> str:
        excel_bytes, stats = run_reconciliation(
            self.state.file_bytes, self.state.config,
            ewb_file_bytes=self.state.ewb_bytes,
            einv_file_bytes=self.state.einv_bytes,
        )
        self.state.excel_bytes = excel_bytes
        self.state.stats       = stats
        return json.dumps(stats, indent=2, ensure_ascii=False)


class ListSheetsTool(BaseTool):
    name: str        = "List Output Sheets"
    description: str = "List sheets in reconciliation output workbook."
    state: object    = None
    class Config: arbitrary_types_allowed = True
    def _run(self, instruction: str = "list") -> str:
        if not self.state.excel_bytes: return "Error: Run reconciliation first."
        return f"Available sheets: {pd.ExcelFile(BytesIO(self.state.excel_bytes)).sheet_names}"


class ReadSheetTool(BaseTool):
    name: str        = "Read Output Sheet"
    description: str = "Read a sheet from the output workbook. Returns up to 100 rows as CSV."
    state: object    = None
    class Config: arbitrary_types_allowed = True
    def _run(self, sheet_name: str) -> str:
        if not self.state.excel_bytes: return "Error: Run reconciliation first."
        try:
            df = pd.read_excel(BytesIO(self.state.excel_bytes), sheet_name=sheet_name)
            return f"Sheet '{sheet_name}' — {len(df)} rows\n\n" + df.head(100).to_csv(index=False)
        except Exception as ex:
            xl = pd.ExcelFile(BytesIO(self.state.excel_bytes))
            return f"Not found. Available: {xl.sheet_names}. Error: {ex}"


class ClassifyBatchTool(BaseTool):
    name: str        = "Classify Invoice Batch for EWB Compliance"
    description: str = "Classify invoices for EWB compliance. Uses threshold confirmed by user."
    state: object    = None
    class Config: arbitrary_types_allowed = True
    def _run(self, invoices_csv: str) -> str:
        try:    df = pd.read_csv(StringIO(invoices_csv))
        except Exception as ex: return f"CSV error: {ex}"
        intra = self.state.config.get("intra_state_threshold", 50000)
        sup   = self.state.config.get("gstin","")
        def fc(*hints):
            for h in hints:
                for c in df.columns:
                    if h.lower() in str(c).lower(): return c
            return None
        val_c = fc("invoice value","grand total","total value","acc value","value")
        to_c  = fc("gstin","gst no","recipient gstin","to gstin")
        inv_c = fc("invoice no","inv no","doc.no","bill no")
        from_c= fc("from gstin","supplier gstin")
        below, genuine, review = [], [], []
        for idx, row in df.iterrows():
            try:
                v  = float(str(row.get(val_c,0)).replace(",","").replace("₹","").strip())
                fg = str(row.get(from_c, sup)) if from_c else sup
                tg = str(row.get(to_c,""))     if to_c   else ""
                if not fg or fg in ("nan",""): fg = sup
                if not tg or tg.upper() in ("N.A","N.A.","NA","NAN",""):
                    below.append(str(row.get(inv_c,f"row_{idx}")) if inv_c else f"row_{idx}")
                    continue
                fs, ts = fg[:2] if len(fg)>=2 else "00", tg[:2] if len(tg)>=2 else "00"
                thresh = 50000 if fs!=ts else intra
                inv_id = str(row.get(inv_c,f"row_{idx}")) if inv_c else f"row_{idx}"
                (below if v < thresh else genuine).append(inv_id)
            except Exception:
                review.append(str(row.get(inv_c,f"row_{idx}")) if inv_c else f"row_{idx}")
        total = len(below)+len(genuine)+len(review)
        return json.dumps({"total_analysed":total,
            "below_threshold_exempt":len(below),
            "genuinely_missing_ewb":len(genuine),
            "needs_manual_review":len(review),
            "genuinely_missing_sample":genuine[:15],
            "thresholds_applied":{"intra_state":f"Rs.{intra:,}","inter_state":"Rs.50,000"},
            "interpretation":(f"Of {total}: {len(below)} exempt, "
                              f"{len(genuine)} genuinely missing EWB, "
                              f"{len(review)} need review.")},indent=2)


class CheckApplicabilityTool(BaseTool):
    name: str        = "Check Single Invoice EWB Applicability"
    description: str = "Check EWB requirement for a specific invoice under Rule 138 CGST Rules 2017."
    state: object    = None
    class Config: arbitrary_types_allowed = True
    def _run(self, from_gstin:str, to_gstin:str, invoice_value:str) -> str:
        try: v = float(str(invoice_value).replace(",","").replace("₹","").strip())
        except: return "Error: invoice_value must be numeric"
        if not to_gstin or to_gstin.upper() in ("N.A","N.A.","NA","","NAN"):
            return (f"POSSIBLE EXPORT — GSTIN is '{to_gstin}'. "
                    f"If export, EWB NOT required under Rule 138(14)(b). "
                    f"Verify DocType. Value: Rs.{v:,.0f}.")
        intra = self.state.config.get("intra_state_threshold", 50000)
        fs    = from_gstin[:2] if len(from_gstin)>=2 else "00"
        ts    = to_gstin[:2]   if len(to_gstin)  >=2 else "00"
        inter = fs != ts
        thresh= 50000 if inter else intra
        label = "inter-state" if inter else f"intra-state (Rs.{intra:,})"
        if v < thresh:
            return f"EWB NOT REQUIRED — {label}. Rs.{v:,.0f} below Rs.{thresh:,.0f}."
        return f"EWB REQUIRED — {label}. Rs.{v:,.0f} exceeds Rs.{thresh:,.0f}."


class RevisedPositionTool(BaseTool):
    name: str        = "Calculate Revised Filing Position"
    description: str = "Recalculate compliance after removing exempt invoices."
    state: object    = None
    class Config: arbitrary_types_allowed = True
    def _run(self, genuinely_missing_ewb_count:str) -> str:
        if not self.state.stats: return "Error: Run reconciliation first."
        try: missing = int(str(genuinely_missing_ewb_count).strip())
        except: return "Error: must be integer string"
        st     = self.state.stats
        total  = st["total_sales"];  ok = st["matched_ok"]
        no_e   = st["sales_no_ewb"]; exempt = max(no_e-missing,0)
        rate   = round((ok+exempt)/total*100,1) if total else 0.0
        new_iss= max(st["issues_count"]-exempt,0)
        intra  = self.state.config.get("intra_state_threshold",50000)
        verdict= ("Ready to file"   if rate>=95 and new_iss<5 else
                  "Review required" if rate>=85 else "Hold — do not file")
        orig_v = ("Hold — do not file" if st["match_rate_pct"]<85 else
                  "Review required"    if st["match_rate_pct"]<95 else "Ready to file")
        return json.dumps({"client_name":st.get("client_name",""),
            "gstin":st.get("gstin",""),"tax_period":st.get("tax_period",""),
            "intra_state_threshold":f"Rs.{intra:,}",
            "total_sales":total,"original_match_rate":st["match_rate_pct"],
            "revised_match_rate":rate,"matched_ok":ok,
            "legitimately_exempt":exempt,"genuinely_missing_ewb":missing,
            "orphan_ewbs":st.get("ewb_no_sales",0),
            "value_mismatches":st.get("val_mismatch",0),
            "einv_irn_missing":st.get("einv_missing_irn",0),
            "original_issues":st["issues_count"],"revised_issues":new_iss,
            "original_verdict":orig_v,"revised_verdict":verdict},indent=2)


def run_gst_analysis(sales_bytes, ewb_bytes, config, einv_bytes=None):
    state = RunState(sales_bytes, ewb_bytes, config, einv_bytes)
    llm   = LLM(model="gemini/gemini-2.5-flash",
                api_key=os.environ.get("GOOGLE_API_KEY",""))

    t_recon = [RunReconciliationTool(state=state)]
    t_inv   = [ListSheetsTool(state=state), ReadSheetTool(state=state),
               ClassifyBatchTool(state=state), CheckApplicabilityTool(state=state)]
    t_comp  = [RevisedPositionTool(state=state)]

    analyst = Agent(role="GST Reconciliation Analyst",
        goal="Run engine and report complete statistics. No interpretation.",
        backstory="Senior GST analyst at a Faridabad CA practice.",
        tools=t_recon, llm=llm, verbose=True, allow_delegation=False, max_iter=3)

    investigator = Agent(role="GST Investigation Specialist",
        goal="Investigate discrepancies. Classify every unmatched invoice.",
        backstory=(f"You investigate before concluding. "
                   f"Intra-state threshold: Rs.{config.get('intra_state_threshold',50000):,}. "
                   f"N.A. GSTIN = possible export."),
        tools=t_inv, llm=llm, verbose=True, allow_delegation=False, max_iter=8)

    officer = Agent(role="GST Compliance Officer",
        goal="Apply GST law. Calculate revised position. Issue defensible verdict.",
        backstory="Deep knowledge of Rule 138 CGST Rules 2017.",
        tools=t_comp, llm=llm, verbose=True, allow_delegation=False, max_iter=3)

    writer = Agent(role="GST Filing Report Writer",
        goal="Write precise filing readiness report.",
        backstory="You write compliance reports for MSME owners and CA teams.",
        tools=[], llm=llm, verbose=True, allow_delegation=False, max_iter=2)

    intra = config.get("intra_state_threshold", 50000)

    tasks = [
        Task(description=(f"Run GSTR-1 reconciliation for {config['client_name']} "
                          f"(GSTIN: {config['gstin']}, Period: {config['tax_period']}). "
                          "Call Run GST Reconciliation. Report all statistics."),
             expected_output="Complete reconciliation statistics.",
             agent=analyst),
        Task(description=(f"Investigate discrepancies. List sheets, read Query Sheet, "
                          f"classify batch, spot-check invoices. "
                          f"Intra-state threshold: Rs.{intra:,}. Inter-state: Rs.50,000. "
                          f"N.A. GSTIN = possible export. Delivery Challans = job work."),
             expected_output="Classification with counts: exempt, missing EWB, exports, review.",
             agent=investigator, context=[]),
        Task(description="Calculate revised filing position using investigation findings. "
                         "Call Calculate Revised Filing Position.",
             expected_output="Revised position with verdict.",
             agent=officer, context=[]),
        Task(description=(f"Write GSTR-1 Filing Readiness Report for "
                          f"{config['client_name']} | {config['gstin']} | {config['tax_period']}. "
                          f"Sections: Original Position, Investigated Position, Revised Position, "
                          f"Action Items, E-Invoice Compliance, Observations. Exact numbers throughout."),
             expected_output="Complete professional filing readiness report.",
             agent=writer, context=[]),
    ]
    tasks[1].context = [tasks[0]]
    tasks[2].context = [tasks[0], tasks[1]]
    tasks[3].context = [tasks[0], tasks[1], tasks[2]]

    crew = Crew(agents=[analyst,investigator,officer,writer],
                tasks=tasks, process=Process.sequential, verbose=True)

    @traceable(name="GSTR-1 Agent Analysis", project_name="gstr1agent",
               tags=["gstr1","crewai","gemini"],
               metadata={"client":config["client_name"],"gstin":config["gstin"],
                         "period":config["tax_period"]})
    def _run():
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:    return crew.kickoff()
        finally: loop.close()

    ctx = contextvars.copy_context()
    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as ex:
        result = ex.submit(ctx.run, _run).result(timeout=600)

    docx_b64 = None
    try:
        from report_generator import generate_word_report
        docx_b64 = base64.b64encode(
            generate_word_report(state.stats, str(result))).decode()
    except Exception as ex:
        print(f"Word report error: {ex}")

    return {"stats": state.stats, "report_text": str(result),
            "excel_b64": base64.b64encode(state.excel_bytes).decode() if state.excel_bytes else None,
            "docx_b64": docx_b64}

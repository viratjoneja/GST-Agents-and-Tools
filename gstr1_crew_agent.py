# GSTR-1 Agentic Reconciliation — CrewAI
# Run in Google Colab. See README.md for setup.
# Cells are demarcated with # ── CELL N ── comments.

# ── CELL 1: Install ──────────────────────────────────────────────────────────
# !pip install -q "crewai[tools]>=0.80.0" anthropic pandas openpyxl xlrd httpx nest_asyncio google-generativeai python-docx

# ── CELL 2: Write recon_engine_api.py ─── see recon_engine_api.py ───────────
# ── CELL 3: Write intelligence.py ──────── see intelligence.py ──────────────

# ── CELL 4: API Key ───────────────────────────────────────────────────────────
# import os
# from google.colab import userdata
# os.environ["GOOGLE_API_KEY"] = userdata.get("GOOGLE_API_KEY")

# ── See full notebook cells in README / Colab notebook ───────────────────────
# The complete runnable notebook is maintained in Google Colab.
# This file captures the agent and tool definitions for version control.

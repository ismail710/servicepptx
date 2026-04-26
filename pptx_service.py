"""
pptx_service.py
---------------
Standalone HTTP service that fills template1.pptx with project data
and returns the result as a base64-encoded PPTX.

Deploy anywhere:
  - Local:      python pptx_service.py
  - Render.com: uses render.yaml in this repo (free tier)
  - Railway:    uses Procfile in this repo (free tier)

POST /generate
  Body (JSON):
    { "project": { ...fields... } }
  Response (JSON):
    { "pptxBase64": "...", "fileName": "PROJECT-001_Title_20260426.pptx" }

GET /health  → 200 { "status": "ok" }
"""

import base64
import copy
import io
import json
import os
import re
import sys
import traceback
from datetime import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer
from pptx import Presentation
from pptx.util import Pt
from lxml import etree

# ---------------------------------------------------------------------------
# Path to the PPTX template (same folder as this script by default)
# ---------------------------------------------------------------------------
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.environ.get("TEMPLATE_PATH", os.path.join(_SCRIPT_DIR, "template1.pptx"))

# ---------------------------------------------------------------------------
# Placeholder → project field mapping
# ---------------------------------------------------------------------------
PLACEHOLDER_MAP = {
    "{{Project_Number}}":     lambda p: p.get("projectNumber", ""),
    "{{Project_Title}}":      lambda p: p.get("projectTitle", ""),
    "{{Project_Lead}}":       lambda p: p.get("projectLead", ""),
    "{{PI_UM6P}}":            lambda p: p.get("piUm6p", ""),
    "{{Starting_Date}}":      lambda p: p.get("startDate", ""),
    "{{Closing_Date}}":       lambda p: p.get("endDate", ""),
    "{{Budget_Amount}}":      lambda p: _fmt_number(p.get("budget", "")),
    "{{Deliverables_List}}":  lambda p: _bullets(p.get("deliverables", p.get("deliverablesCsv", ""))),
    "{{Risk_Alert}}":         lambda p: _bullets(p.get("risks_and_alerts", p.get("risksAndAlertsCsv", ""))),
    "{{Tasks_Done}}":         lambda p: _bullets(p.get("done", p.get("doneCsv", ""))),
    "{{Tasks_InProgress}}":   lambda p: _bullets(p.get("in_progress", p.get("inProgressCsv", ""))),
    "{{Tasks_NextSteps}}":    lambda p: _bullets(p.get("planned", p.get("plannedCsv", ""))),
    "{{CAPEX_Commitments}}":  lambda p: _fmt_number(p.get("capex_commitments", p.get("capexCommitments", ""))),
    "{{CAPEX_Expenses}}":     lambda p: _fmt_number(p.get("capex_expenses", p.get("capexExpenses", ""))),
    "{{OPEX_Commitments}}":   lambda p: _fmt_number(p.get("opex_commitments", p.get("opexCommitments", ""))),
    "{{OPEX_Expenses}}":      lambda p: _fmt_number(p.get("opex_expenses", p.get("opexExpenses", ""))),
}


def _fmt_number(val: str) -> str:
    """Format numeric strings with thousands separator (e.g. 250000 → 250 000)."""
    v = str(val).strip()
    try:
        n = int(float(v))
        return f"{n:,}".replace(",", " ")
    except (ValueError, TypeError):
        return v


def _bullets(csv_val: str) -> str:
    """Convert pipe-separated CSV to bullet lines: 'A | B' → '• A\n• B'."""
    if not csv_val:
        return ""
    items = [x.strip() for x in str(csv_val).split("|") if x.strip()]
    return "\n".join(f"• {item}" for item in items)


# ---------------------------------------------------------------------------
# Core PPTX filling logic
# ---------------------------------------------------------------------------

def _replace_in_paragraph(para, replacements: dict):
    """
    Replace all {{TOKEN}} occurrences in a paragraph.
    Handles tokens that are split across multiple runs by first merging
    all run text, performing the replacement, then writing back into the
    first run and clearing the rest.
    """
    if not para.runs:
        return

    # Merge all run text
    full_text = "".join(r.text for r in para.runs)

    # Check if any replacement is needed
    if not any(token in full_text for token in replacements):
        return

    # Perform all replacements on the merged text
    new_text = full_text
    for token, value in replacements.items():
        new_text = new_text.replace(token, value)

    # Write back: put everything in the first run, clear subsequent runs
    para.runs[0].text = new_text
    for run in para.runs[1:]:
        run.text = ""


def fill_pptx_base64(project_data: dict) -> dict:
    """
    Load template1.pptx, replace all placeholders, and return
    { "pptxBase64": "...", "fileName": "..." }
    """
    prs = Presentation(TEMPLATE_PATH)

    # Build token → value dict
    replacements = {
        token: fn(project_data)
        for token, fn in PLACEHOLDER_MAP.items()
    }

    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                _replace_in_paragraph(para, replacements)

    # Serialize to bytes
    buf = io.BytesIO()
    prs.save(buf)
    pptx_bytes = buf.getvalue()

    # Build file name
    title_slug = re.sub(r"[^A-Za-z0-9_]", "_",
                        project_data.get("projectTitle", "Projet"))[:40]
    proj_num   = re.sub(r"[^A-Za-z0-9_-]", "_",
                        project_data.get("projectNumber", "XXX"))
    date_str   = datetime.now().strftime("%Y%m%d")
    file_name  = f"{proj_num}_{title_slug}_{date_str}.pptx"

    return {
        "pptxBase64": base64.b64encode(pptx_bytes).decode("utf-8"),
        "fileName":   file_name,
    }


# ---------------------------------------------------------------------------
# HTTP request handler
# ---------------------------------------------------------------------------

class PPTXHandler(BaseHTTPRequestHandler):

    def log_message(self, fmt, *args):  # suppress default noisy logging
        print(f"[{self.address_string()}] {fmt % args}", file=sys.stderr)

    def _send_json(self, status: int, data: dict):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        if self.path.rstrip("/") in ("/health", "/api/health"):
            self._send_json(200, {"status": "ok", "template": TEMPLATE_PATH})
        else:
            self._send_json(404, {"error": "Not found"})

    def do_POST(self):
        if self.path.rstrip("/") not in ("/generate", "/api/generate"):
            self._send_json(404, {"error": "Not found"})
            return

        try:
            length  = int(self.headers.get("Content-Length", 0))
            raw     = self.rfile.read(length) if length else b"{}"
            payload = json.loads(raw.decode("utf-8"))
        except Exception as exc:
            self._send_json(400, {"error": f"Invalid JSON: {exc}"})
            return

        # Accept both { project: {...} } and flat { projectTitle: ... }
        project_data = payload.get("project", payload)

        try:
            result = fill_pptx_base64(project_data)
            self._send_json(200, result)
        except Exception as exc:
            traceback.print_exc()
            self._send_json(500, {"error": str(exc)})


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def run_server(port: int = 7771):
    if not os.path.exists(TEMPLATE_PATH):
        print(f"ERROR: template not found at {TEMPLATE_PATH}", file=sys.stderr)
        sys.exit(1)
    server = HTTPServer(("0.0.0.0", port), PPTXHandler)
    print(f"PPTX service running on port {port} — POST /generate  GET /health")
    print(f"Template: {TEMPLATE_PATH}")
    server.serve_forever()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 7771))
    run_server(port)

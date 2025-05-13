from http.server import BaseHTTPRequestHandler
import json
import re

def _normalize(text: str) -> str:
    """Lower‑case & collapse whitespace for fuzzy matching."""
    return re.sub(r"\s+", " ", text or "").lower().strip()

# ------------------------------------------------------------------ #
# 1) Expert playbook of Excel/VBA scenarios
# ------------------------------------------------------------------ #
SCENARIOS = {
    "pivot": [
        {
            "match": ["pivot", "calculated field", "grouping", "pivot table"],
            "title": "Use PivotCache to refresh without flicker",
            "description": (
                "If you're looping over PivotTables or refreshing them frequently, "
                "use the .PivotCache.Refresh method to reduce lag."
            ),
            "reference": None,
            "tip": "Dim pc As PivotCache: Set pc = pt.PivotCache: pc.Refresh"
        },
        {
            "match": ["field", "calculated field", "formula"],
            "title": "Add calculated fields programmatically",
            "description": (
                "You can insert calculated fields directly via VBA using .CalculatedFields.Add."
            ),
            "reference": None,
            "tip": "pt.CalculatedFields.Add \"Margin\", \"=Revenue - Cost\", True"
        }
    ],
    {
    # ----------  USERFORMS  ----------
    "userform": [
        {
            "match": ["multipage", "tabs", "multi page"],
            "title": "Use MultiPage control for tabbed sections",
            "description": (
                "MultiPage lets you group related inputs into tabs and keeps forms compact."
            ),
            "reference": "Microsoft Forms 2.0 Object Library",
            "tip": "Insert ➜ ActiveX ➜ MultiPage  •  In code: With Me.MultiPage1.Pages(0)..."
        },
        {
            "match": ["calendar", "date", "monthview"],
            "title": "MonthView control for reliable date picking",
            "description": (
                "MonthView offers an intuitive calendar UI and avoids manual date validation."
            ),
            "reference": "Microsoft Windows Common Controls‑2 6.0 (MSCOMCT2.OCX)",
            "tip": "If MonthView isn't listed, register MSCOMCT2.OCX then enable via Tools ➜ Additional Controls."
        },
        {
            "match": ["events", "broadcast", "notify", "scripting.dictionary"],
            "title": "Broadcast custom events with WithEvents Dictionary",
            "description": (
                "Combine `WithEvents` and `Scripting.Dictionary` to raise events from models "
                "and have multiple UserForms listen without tight coupling."
            ),
            "reference": "Microsoft Scripting Runtime",
            "tip": "Private WithEvents mBus As Scripting.Dictionary  ' ... raise mBus(\"Change\") = True"
        }
    ],
    # ----------  PERFORMANCE  ----------
    "performance": [
        {
            "match": ["loops", "long", "recalc"],
            "title": "Manual calculation mode inside long loops",
            "description": (
                "Set `Application.Calculation = xlCalculationManual` before bulk updates "
                "then restore to prevent Excel recalculating after every cell write."
            ),
            "reference": None,
            "tip": "Application.Calculation = xlCalculationManual  ' ...code...  Application.CalculateFull"
        },
        {
            "match": ["screen", "flicker"],
            "title": "Turn off ScreenUpdating & PageBreaks for flicker‑free speed",
            "description": (
                "Disabling these two flags can reduce UI churn by 90 % on big sheets."
            ),
            "reference": None,
            "tip": "With Application: .ScreenUpdating=False: ActiveSheet.DisplayPageBreaks=False: End With"
        }
    ],
    # ----------  CHARTS  ----------
    "charts": [
        {
            "match": ["export", "pdf", "print"],
            "title": "One‑line PDF export with Chart.ExportAsFixedFormat",
            "description": (
                "Export any chart to PDF/XPS without touching `Print` dialogs."
            ),
            "reference": None,
            "tip": "ActiveChart.ExportAsFixedFormat xlTypePDF, \"C:\\Temp\\Chart.pdf\""
        },
        {
            "match": ["animate", "ani", "smooth"],
            "title": "Flicker‑free animation via hidden XLM DISPOFF/DISPON",
            "description": (
                "Call legacy XLM commands to suppress redraw between frame updates."
            ),
            "reference": None,
            "tip": "Application.ExecuteExcel4Macro \"DISPOFF\" : '...update series...' : Application.ExecuteExcel4Macro \"DISPON\""
        }
    ],
}

# ------------------------------------------------------------------ #
# 2) Core handler
# ------------------------------------------------------------------ #
class handler(BaseHTTPRequestHandler):
    def _json_response(self, data, status=200):
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(json.dumps(data, indent=2).encode())

    # ---------- POST (primary) ----------
    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        raw    = self.rfile.read(length or 0)
        try:
            req = json.loads(raw or "{}")
        except json.JSONDecodeError:
            return self._json_response({"error": "Invalid JSON body"}, 400)

        ctx   = _normalize(req.get("context", ""))
        stype = _normalize(req.get("type", "general"))
        lang  = _normalize(req.get("language", "vba"))

        matches = []

        # 1) direct type match
        for rec in SCENARIOS.get(stype, []):
            if any(k in ctx for k in rec["match"]):
                matches.append(rec)

        # 2) fallback: scan all scenarios if nothing matched yet
        if not matches:
            for topic, rules in SCENARIOS.items():
                for rec in rules:
                    if any(k in ctx for k in rec["match"]):
                        matches.append(rec)
                        break

        # 3) final fallback
        if not matches:
            matches.append({
                "title": "No specific pro tips found",
                "description": "Try adding more detail or keywords (e.g. 'MultiPage', 'MonthView', 'animation').",
                "reference": None,
                "tip": None
            })

        self._json_response({
            "topic": stype,
            "recommendations": matches[: req.get("maxItems", 5)]
        })

    # ---------- GET (nice for browsers) ----------
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain")
        self.end_headers()
        msg = (
            "ExcelGPT Contextual Features API\n"
            "POST JSON: {context:'text', type:'userform|charts|performance', language:'vba|script'}"
        )
        self.wfile.write(msg.encode())

from http.server import BaseHTTPRequestHandler
import json, re

# ────────────────────────────────────────────────────────────────
def _normalize(text: str) -> str:
    """lower‑case & collapse whitespace for fuzzy matching"""
    return re.sub(r"\s+", " ", text or "").lower().strip()


# ----------------------------------------------------------------
# 1) Expert playbook of Excel/VBA scenarios
# ----------------------------------------------------------------
SCENARIOS = {
    "pivot": [
        {
            "match": ["pivot", "calculated field", "grouping", "pivot table"],
            "title": "Use PivotCache to refresh without flicker",
            "description": (
                "When looping or frequent refreshes, call .PivotCache.Refresh "
                "instead of pt.RefreshTable to avoid UI lag."
            ),
            "reference": None,
            "tip": "Dim pc As PivotCache: Set pc = pt.PivotCache: pc.Refresh"
        },
        {
            "match": ["field", "calculated field", "formula"],
            "title": "Insert calculated fields programmatically",
            "description": (
                "Use .CalculatedFields.Add to add or update formulas in code."
            ),
            "reference": None,
            "tip": 'pt.CalculatedFields.Add "Margin", "=Revenue-Cost", True'
        }
    ],

    # ----------  USERFORMS  ----------
    "userform": [
        {
            "match": ["multipage", "tabs", "multi page"],
            "title": "MultiPage for tabbed sections",
            "description": "Group related inputs into tabs to keep large forms compact.",
            "reference": "Microsoft Forms 2.0 Object Library",
            "tip": "Insert → ActiveX → MultiPage"
        },
        {
            "match": ["calendar", "date", "monthview"],
            "title": "MonthView control for date picking",
            "description": "Native calendar UI avoids manual date validation.",
            "reference": "MSCOMCT2.OCX",
            "tip": "Register MSCOMCT2, then Tools → Additional Controls → MonthView"
        },
        {
            "match": ["events", "broadcast", "notify", "scripting.dictionary"],
            "title": "Broadcast custom events via WithEvents Dictionary",
            "description": "Raise events from models and let multiple forms listen loosely.",
            "reference": "Microsoft Scripting Runtime",
            "tip": "Private WithEvents mBus As Scripting.Dictionary"
        }
    ],

    # ----------  PERFORMANCE  ----------
    "performance": [
        {
            "match": ["loops", "long", "recalc"],
            "title": "Manual calculation during long loops",
            "description": "Set Application.Calculation = xlCalculationManual, restore after.",
            "reference": None,
            "tip": "Application.Calculation = xlCalculationManual"
        },
        {
            "match": ["screen", "flicker"],
            "title": "Disable ScreenUpdating & PageBreaks",
            "description": "Reduces UI churn by ~90 % on big sheets.",
            "reference": None,
            "tip": "With Application: .ScreenUpdating=False: ActiveSheet.DisplayPageBreaks=False"
        }
    ],

    # ----------  CHARTS  ----------
    "charts": [
        {
            "match": ["export", "pdf", "print"],
            "title": "One‑line PDF export",
            "description": "Use Chart.ExportAsFixedFormat to skip the Print dialog.",
            "reference": None,
            "tip": "ActiveChart.ExportAsFixedFormat xlTypePDF, \"C:\\Temp\\Chart.pdf\""
        },
        {
            "match": ["animate", "ani", "smooth"],
            "title": "Flicker‑free animation (DISPOFF/DISPON)",
            "description": "Call legacy XLM to suppress redraw between frames.",
            "reference": None,
            "tip": "Application.ExecuteExcel4Macro \"DISPOFF\" : … : Application.ExecuteExcel4Macro \"DISPON\""
        }
    ]
}

# ----------------------------------------------------------------
# 2) Generic keyword → hint map (fallback)
# ----------------------------------------------------------------
GENERIC_HINTS = {
    "pivot":      "Use PivotCache.Refresh and CalculatedFields.Add for dynamic fields.",
    "userform":   "Try MultiPage for tabs and MonthView for date inputs.",
    "chart":      "Chart.ExportAsFixedFormat makes instant PDFs; DISPOFF cuts flicker.",
    "performance":"Turn off ScreenUpdating and switch calc to Manual inside loops."
}

# ----------------------------------------------------------------
# 3) HTTP handler
# ----------------------------------------------------------------
class handler(BaseHTTPRequestHandler):
    def _json(self, obj, status=200):
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(json.dumps(obj, indent=2).encode())

    # ---------- POST ----------
    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        data   = json.loads(self.rfile.read(length) or "{}")

        ctx   = _normalize(data.get("context", ""))
        stype = _normalize(data.get("type", "general"))

        matches = [
            rec for rec in SCENARIOS.get(stype, [])
            if any(k in ctx for k in rec["match"])
        ]

        # fallback scan across all topics
        if not matches:
            for rules in SCENARIOS.values():
                for rec in rules:
                    if any(k in ctx for k in rec["match"]):
                        matches.append(rec)
                        break
                if matches: break

        # generic hint fallback
        if not matches:
            for kw, hint in GENERIC_HINTS.items():
                if kw in ctx:
                    matches.append({
                        "title": f"General tip for '{kw}' tasks",
                        "description": hint,
                        "reference": None,
                        "tip": None
                    })
                    break

        # final "need more detail"
        if not matches:
            matches.append({
                "title": "Need more detail",
                "description": "Add keywords like 'pivot', 'chart', 'userform', or describe the control you’re using.",
                "reference": None,
                "tip": None
            })

        self._json({
            "topic": stype,
            "recommendations": matches[: data.get("maxItems", 5)]
        })

    # ---------- GET ----------
    def do_GET(self):
        self._json({"hint": "POST JSON {context:'build XYZ', type:'charts'...} for recommendations"})

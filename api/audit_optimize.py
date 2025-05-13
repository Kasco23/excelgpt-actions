from http.server import BaseHTTPRequestHandler
import json, re, textwrap

# ----------------------------------------------------------------
# 1) LINT RULES  (regex → message & fix)
# ----------------------------------------------------------------
RULES = [
    # implicit variable (Option Explicit missing OR undeclared dim)
    {
        "pattern": r"\bFor\s+(\w+)\s*=",
        "message": "Loop variable '{0}' is implicit. Add Dim {0} As Long.",
        "fix": "Add “Option Explicit” and declare variables."
    },
    {
        "pattern": r"\bCells?\(.+?\)\s*=",
        "message": "Cell‑by‑cell write. Buffer Range to Variant.",
        "fix": "Read Range.Value2 to Variant, process in memory, write back once."
    },
    {
        "pattern": r"\.Select\b|\bSelection\.",
        "message": "Use of .Select / Selection slows code.",
        "fix": "Work directly with Range objects (e.g. ws.Range(…))."
    },
    {
        "pattern": r"Application\.ScreenUpdating\s*=\s*True",
        "message": "ScreenUpdating turned on mid‑macro.",
        "fix": "Only re‑enable at the very end."
    },
    {
        "pattern": r"On\s+Error\s+Resume\s+Next",
        "message": "Blind error suppression.",
        "fix": "Use structured handler or test Err.Number."
    },
    {
        "pattern": r"\bDo\s+While\s+Not\s+(.+?)\.EOF",
        "message": "DAO/ADODB recordset loop; consider GetRows for speed.",
        "fix": "Use rs.GetRows to bulk‑load to array."
    },
    {
        "pattern": r"\bApplication\.Calculate\b",
        "message": "Full recalculation call; consider CalculateFullRebuild only if necessary.",
        "fix": "Limit to affected Range.Calculate where possible."
    },
    {
        "pattern": r"ActiveWorkbook\.Save",
        "message": "Saving workbook mid‑macro can freeze UI.",
        "fix": "Buffer changes; save once at end or use AutoSave off‑peak."
    },
    {
        "pattern": r"DoEvents\(\)",
        "message": "Frequent DoEvents in tight loops slows performance.",
        "fix": "Throttle with counter, or remove when batch processing."
    },
    {
        "pattern": r"\bVariant\b\s*=",
        "message": "Variant used explicitly; specify concrete type for speed.",
        "fix": "Use Long, Double, String, etc."
    },
    {
        "pattern": r"\bWorksheetFunction\.([A-Za-z]+)\(",
        "message": "WorksheetFunction call in loops is slow.",
        "fix": "Port math into VBA or call once on whole range."
    },
    {
        "pattern": r"\bVLookup\(",
        "message": "VLookup in VBA loops is slow.",
        "fix": "Use Dictionary or INDEX/MATCH on entire arrays."
    },
    {
        "pattern": r"^\s*Public\s+",
        "flags": re.MULTILINE,
        "message": "Public scope—review necessity; many modules don’t need it.",
        "fix": "Change to Private unless accessed externally."
    },
    {
        "pattern": r"CreateObject\(\"Scripting\.Dictionary\"",
        "message": "Dictionary created late‑bound—use early binding.",
        "fix": "Add reference 'Microsoft Scripting Runtime' and use New Scripting.Dictionary."
    }
]

# ----------------------------------------------------------------
# 2) Handler
# ----------------------------------------------------------------
class handler(BaseHTTPRequestHandler):
    def _json(self, obj, status=200):
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(json.dumps(obj, indent=2).encode())

    # ---------- POST ----------
    def do_POST(self):
        raw = self.rfile.read(int(self.headers.get("Content-Length", 0)) or 0)
        try:
            body = json.loads(raw or "{}")
            code = body["code"]  # original VBA text
        except Exception:
            return self._json({"error": "POST JSON {code:'VBA text'}"}, 400)

        lines   = code.splitlines()
        issues  = []
        flagged = set()

        # run rules
        for ln, line in enumerate(lines, start=1):
            for rule in RULES:
                flags = rule.get("flags", 0)
                m = re.search(rule["pattern"], line, flags)
                if m:
                    key = (rule["pattern"], ln)   # avoid duplicates same line
                    if key in flagged:
                        continue
                    flagged.add(key)
                    msg = rule["message"].format(*m.groups())
                    issues.append({
                        "line": ln,
                        "message": msg,
                        "suggestion": rule["fix"]
                    })
                    lines[ln-1] += f"    ' 🔍 ISSUE: {msg}"
                    break

        # build checklist (top 5 unique suggestions)
        checklist = []
        seen_fix  = set()
        for iss in issues:
            if iss["suggestion"] and iss["suggestion"] not in seen_fix:
                checklist.append("• " + iss["suggestion"])
                seen_fix.add(iss["suggestion"])
            if len(checklist) >= 5:
                break

        self._json({
            "annotatedCode": "\n".join(lines),
            "findings": issues,
            "todo": checklist or ["Looks solid! No common issues found."]
        })

    # ---------- GET ----------
    def do_GET(self):
        self._json({"hint": "POST JSON {code:'<paste VBA code>'} to audit"})

from http.server import BaseHTTPRequestHandler
import json, base64, io, zipfile, re
from collections import defaultdict

# ----- Regexes for code analysis -----
PROC_RE   = re.compile(r"^\s*(Public |Private )?(Sub|Function)\s+([A-Za-z0-9_]+)", re.I | re.M)
CALL_RE   = re.compile(r"\bCall\s+([A-Za-z0-9_]+)|\b([A-Za-z0-9_]+)\s*\(", re.I)
RANGE_RE  = re.compile(r"\bRange\(\"([A-Z0-9:]+)\"\)|Cells?\(", re.I)
CTRL_RE   = re.compile(r"\.(Top|Left|Visible|Enabled|Caption)\b", re.I)
ERROR_RE  = re.compile(r"On\s+Error\s+(Resume\s+Next|GoTo)", re.I)

def parse_module(name, text, graph, effects, hotspots):
    funcs = {m.group(3).lower(): m.start() for m in PROC_RE.finditer(text)}
    for fname, start in funcs.items():
        block = text[start:]
        end   = block.find("End Sub")
        if end == -1: end = block.find("End Function")
        block = block[: end if end != -1 else None]

        for c1, c2 in CALL_RE.findall(block):
            called = (c1 or c2).strip().lower()
            if called and called not in funcs:
                graph[fname].add(called)

        if RANGE_RE.search(block): effects[fname].add("Reads/writes cells")
        if CTRL_RE.search(block): effects[fname].add("Touches form controls")
        if "FileSystemObject" in block: effects[fname].add("File system access")
        if ERROR_RE.search(block): hotspots.add(fname)

class handler(BaseHTTPRequestHandler):
    def _json(self, obj, status=200):
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(json.dumps(obj, indent=2).encode())

    def do_POST(self):
        content_type = self.headers.get("Content-Type", "").lower()
        zip_bytes = None

        # --- Option 1: zipBase64 via JSON
        if "application/json" in content_type:
            try:
                data = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0)) or 0))
                if "zipBase64" not in data:
                    raise ValueError("Missing 'zipBase64'")
                zip_bytes = base64.b64decode(data["zipBase64"])
            except Exception:
                return self._json({"error": "Invalid JSON with 'zipBase64'"}, 400)

        # --- Option 2: raw ZIP file upload
        elif "application/zip" in content_type:
            zip_bytes = self.rfile.read(int(self.headers.get("Content-Length", 0)) or 0)

        else:
            return self._json({"error": "Expected zipBase64 (JSON) or ZIP file upload"}, 400)

        try:
            zf = zipfile.ZipFile(io.BytesIO(zip_bytes))
        except Exception:
            return self._json({"error": "Invalid ZIP format"}, 400)

        # --- Read .bas modules from zip ---
        modules = {}
        for name in zf.namelist():
            if name.lower().endswith(".bas"):
                modules[name] = zf.read(name).decode("utf-8", "ignore")

        if not modules:
            return self._json({"error": "No .bas files found in ZIP"}, 400)

        # --- Analyze ---
        graph    = defaultdict(set)
        effects  = defaultdict(set)
        hotspots = set()

        for mname, mtext in modules.items():
            parse_module(mname, mtext, graph, effects, hotspots)

        story = []
        for proc, calls in graph.items():
            story.append(f"{proc}() ➜ calls: {', '.join(calls) or '—'}")
            if effects[proc]:
                story.append(f"    side-effects: {', '.join(effects[proc])}")
            if proc in hotspots:
                story.append("    ⚠️ On Error handler detected — review this block.")

        self._json({
            "callGraph": {k: list(v) for k, v in graph.items()},
            "effects": {k: list(v) for k, v in effects.items()},
            "errorHotspots": list(hotspots),
            "storyboard": "\n".join(story)
        })

    def do_GET(self):
        self._json({
            "hint": "POST zipBase64 as JSON or upload ZIP (Content-Type: application/zip)"
        })

from http.server import BaseHTTPRequestHandler
import json, re, base64, io, zipfile, textwrap
from collections import defaultdict

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
        if end == -1:
            end = block.find("End Function")
        block = block[: end if end != -1 else None]

        # call map
        for c1, c2 in CALL_RE.findall(block):
            called = (c1 or c2).strip().lower()
            if called and called not in funcs:
                graph[fname].add(called)

        # side‑effects
        if RANGE_RE.search(block):
            effects[fname].add("Writes/reads cell ranges")
        if CTRL_RE.search(block):
            effects[fname].add("Touches form/ActiveX controls")
        if "FileSystemObject" in block:
            effects[fname].add("File system operations")
        if ERROR_RE.search(block):
            hotspots.add(fname)

class handler(BaseHTTPRequestHandler):
    def _json(self, obj, status=200):
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(json.dumps(obj, indent=2).encode())

    # ---------- POST ----------
    def do_POST(self):
        body = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0)) or 0))

        modules = {}   # {moduleName: codeText}

        # 1) single .bas text
        if "basText" in body:
            modules["Module1"] = body["basText"]

        # 2) base‑64 zip
        elif "zipBase64" in body:
            try:
                zbytes = base64.b64decode(body["zipBase64"])
                zf     = zipfile.ZipFile(io.BytesIO(zbytes))
                for name in zf.namelist():
                    if name.lower().endswith(".bas"):
                        modules[name] = zf.read(name).decode("utf-8", "ignore")
            except Exception:
                return self._json({"error": "Bad base64 ZIP upload"}, 400)

        else:
            return self._json({"error": "Provide basText or zipBase64"}, 400)

        # ------------- analysis -------------
        graph    = defaultdict(set)
        effects  = defaultdict(set)
        hotspots = set()

        for mname, mtext in modules.items():
            parse_module(mname, mtext, graph, effects, hotspots)

        # build storyboard
        story = []
        for proc, calls in graph.items():
            story.append(f"{proc}()  ➜  calls: {', '.join(calls) or '—'}")
            if effects[proc]:
                story.append(f"    side‑effects: {', '.join(effects[proc])}")
            if proc in hotspots:
                story.append("    ⚠️  has raw On Error — review handling")

        self._json({
            "callGraph": {k: list(v) for k, v in graph.items()},
            "effects": {k: list(v) for k, v in effects.items()},
            "errorHotspots": list(hotspots),
            "storyboard": "\n".join(story) or "No procedures found."
        })

    # ---------- GET ----------
    def do_GET(self):
        self._json({"hint": "POST {basText:'...'} or {zipBase64:'...'}"})

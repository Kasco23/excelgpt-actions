from http.server import BaseHTTPRequestHandler
import json, textwrap, datetime

def _dbg(label: str) -> str:
    return f'    Debug.Print "{label}"; Timer'

class handler(BaseHTTPRequestHandler):
    def _json(self, obj, status=200):
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(json.dumps(obj, indent=2).encode())

    # ---------- POST ----------
    def do_POST(self):
        req = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0)) or 0))

        name   = req["moduleName"]
        mtype  = req.get("moduleType", "helper").lower()
        steps  = req.get("steps", [])
        hdr    = req.get("headerComment", "")
        dbgOn  = req.get("includeDebugPrint", False)
        errOn  = req.get("includeErrorHandler", True)

        stamp  = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
        header = f"'{hdr}\n'Generated: {stamp}\n\n" if hdr else ""
        out    = [header + "Option Explicit\n"]

        # ------------ templates ------------
        if mtype == "sequence":
            out.append(f"'=====  SEQUENCE MODULE: {name}  ======\n")
            for s in steps:
                proc = s.replace(" ", "")
                out.append(textwrap.dedent(f"""\
                    Public Sub {proc}()
                        {_dbg(s) if dbgOn else ''}
                        ' TODO: {s.lower()}
                    End Sub\n"""))
            out.append("Public Sub Run()\n")
            for s in steps: out.append(f"    {s.replace(' ', '')}\n")
            out.append("End Sub\n")

        elif mtype == "settings":
            out.append("'=====  SETTINGS MODULE  ======\n")
            out.append("Public Const APP_TITLE As String = \"MyWorkbook\"\n")

        elif mtype == "error":
            out.append("'=====  CENTRAL ERROR HANDLER  ======\n")
            out.append(textwrap.dedent("""\
                Public Sub HandleError(ByVal proc As String, ByVal errObj As ErrObject)
                    Debug.Print \"[ERROR]\", proc, errObj.Number, errObj.Description
                    MsgBox \"Error \" & errObj.Number & \" in \" & proc & vbCrLf & errObj.Description, vbExclamation
                End Sub\n"""))

        elif mtype == "events":
            out.append("'=====  EVENT STUBS  ======\n")
            out.append(textwrap.dedent("""\
                Public Sub wb_Open():  'Workbook_Open stub : End Sub
                Public Sub ws_Change(ByVal Target As Range):  'Worksheet_Change stub : End Sub\n"""))

        else:  # helper
            out.append(f"'=====  HELPER MODULE: {name}  ======\n")
            out.append(textwrap.dedent(f"""\
                Public Function {name}_Helper() As Boolean
                    On Error GoTo ErrHandler
                    {_dbg('start') if dbgOn else ''}

                    ' …logic…

                    {_dbg('end') if dbgOn else ''}
                    {name}_Helper = True
                Done:
                    Exit Function
                ErrHandler:
                    {'HandleError \"' + name + '_Helper\", Err' if errOn else 'MsgBox Err.Description'}
                    {name}_Helper = False
                    Resume Done
                End Function\n"""))

        # ------------ guidance ------------
        todo = []
        if mtype == "sequence":
            todo += [
                f"• Call {name}.Run from a controller sub or Workbook_Open.",
                "• Replace each step stub with real code."
            ]
        if errOn and mtype != "error":
            todo.append("• Add or verify a central error‑handler module exists (moduleType='error').")
        if dbgOn:
            todo.append("• Open Immediate Window (Ctrl+G) to watch Debug.Print output.")
        if not todo:
            todo.append("• Insert this module and hook it where needed.")

        self._json({"code": "\n".join(out), "todo": todo})

    # ---------- GET ----------
    def do_GET(self):
        self._json({"hint": "POST JSON {moduleName:'MyModule'} to generate boilerplate"})

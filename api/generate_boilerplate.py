from http.server import BaseHTTPRequestHandler
import json, textwrap, datetime

# ────────────────────────────────────────────────────────────────
# Helper: create a Debug.Print line if requested
def _dbg(label: str) -> str:
    return f'    Debug.Print "{label}"; Timer'

# ────────────────────────────────────────────────────────────────
class handler(BaseHTTPRequestHandler):
    # quick JSON response wrapper
    def _respond(self, obj, status=200):
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.end_headers()
        self.wfile.write(json.dumps(obj, indent=2).encode())

    # ------------------------------------------------------------------ #
    # POST  → generate module
    # ------------------------------------------------------------------ #
    def do_POST(self):
        try:
            raw = self.rfile.read(int(self.headers.get("Content-Length", 0)) or 0)
            req = json.loads(raw or "{}")
        except Exception:
            return self._respond({"error": "Invalid JSON body"}, 400)

        # ------------ parameters
        name   = req["moduleName"]
        mtype  = req.get("moduleType", "helper").lower()
        steps  = req.get("steps", [])
        hdrTxt = req.get("headerComment", "")
        dbgOn  = req.get("includeDebugPrint", False)
        errOn  = req.get("includeErrorHandler", True)

        # ------------ header comment
        stamp   = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M UTC")
        header  = f"'{hdrTxt}\n'Generated: {stamp}\n\n" if hdrTxt else ""
        code    = [header + "Option Explicit\n"]

        # ----------------------------------------------------------------
        # TEMPLATES by module type
        # ----------------------------------------------------------------
        if mtype == "sequence":
            code.append(f"'=====  SEQUENCE MODULE: {name}  ======\n")
            for step in steps:
                proc = step.replace(" ", "")
                code.append(textwrap.dedent(f"""\
                    Public Sub {proc}()
                        {_dbg(step) if dbgOn else ''}
                        ' TODO: implement {step.lower()}
                    End Sub\n"""))
            code.append("Public Sub Run()\n")
            for step in steps:
                code.append(f"    {step.replace(' ', '')}\n")
            code.append("End Sub\n")

        elif mtype == "settings":
            code.append("'=====  SETTINGS (constants / enums)  ======\n")
            code.append("Public Const APP_TITLE As String = \"MyWorkbook\"\n")
            code.append("'Add further constants below …\n")

        elif mtype == "error":
            code.append("'=====  CENTRAL ERROR HANDLER  ======\n")
            code.append(textwrap.dedent("""\
                Public Sub HandleError(ByVal proc As String, ByVal errObj As ErrObject)
                    Debug.Print \"[ERROR]\", proc, errObj.Number, errObj.Description
                    MsgBox \"Error \" & errObj.Number & \" in \" & proc & vbCrLf & errObj.Description, vbExclamation
                End Sub\n"""))

        elif mtype == "events":
            code.append("'=====  WORKBOOK / SHEET EVENT STUBS  ======\n")
            code.append(textwrap.dedent("""\
                Public Sub wb_Open()
                    'Workbook_Open stub
                End Sub

                Public Sub ws_Change(ByVal Target As Range)
                    'Worksheet_Change stub
                End Sub\n"""))

        else:  # generic helper
            code.append(f"'=====  HELPER MODULE: {name}  ======\n")
            code.append(textwrap.dedent(f"""\
                Public Function {name}_Helper() As Boolean
                    On Error GoTo ErrHandler
                    {_dbg('Start helper') if dbgOn else ''}

                    ' …logic…

                    {_dbg('End helper') if dbgOn else ''}
                    {name}_Helper = True
                Done:
                    Exit Function
                ErrHandler:
                    {'HandleError \"' + name + '_Helper\", Err' if errOn else 'MsgBox Err.Description'}
                    {name}_Helper = False
                    Resume Done
                End Function\n"""))

        # ----------------------------------------------------------------
        # TODO checklist (guidance, non‑project‑specific)
        # ----------------------------------------------------------------
        todo = []
        if mtype == "sequence":
            todo.append(f"• Call {name}.Run from your main controller or Workbook_Open.")
            todo.append("• Flesh out each step stub with real code.")
        if errOn and mtype != "error":
            todo.append("• Ensure a central error‑handler module exists (moduleType='error').")
        if dbgOn:
            todo.append("• Use Ctrl+G to watch Debug.Print output during testing.")

        self._respond({"code": "\n".join(code), "todo": todo})

    # ------------------------------------------------------------------ #
    # GET  → simple help text
    # ------------------------------------------------------------------ #
    def do_GET(self):
        self._respond({"hint": "POST JSON to generate a VBA module. Required: moduleName"})

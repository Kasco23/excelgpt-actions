from http.server import BaseHTTPRequestHandler
import json, textwrap

class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        length = int(self.headers['Content-Length'])
        body   = json.loads(self.rfile.read(length))
        name   = body['moduleName']
        hdr    = body.get('headerComment', '')
        err_on = body.get('includeErrorHandler', True)

        header = f"'{hdr}\n" if hdr else ''
        code   = textwrap.dedent(f"""{header}Option Explicit

Public Sub {name}_Main()
    {"On Error GoTo ErrHandler" if err_on else "' add error handling if needed"}
    ' TODO: Your code here

Done:
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation, "Error in {name}_Main"
    Resume Done
End Sub
""")
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps({'code': code}).encode())

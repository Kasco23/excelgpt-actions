from http.server import BaseHTTPRequestHandler
import json

class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        response = {
            "code": "Option Explicit\n\nPublic Sub Test()\n    MsgBox \"Hello\"\nEnd Sub"
        }
        self.wfile.write(json.dumps(response).encode())

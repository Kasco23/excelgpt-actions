from http.server import BaseHTTPRequestHandler
import json

class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        response = {
            "gems": [
                "Hidden Chart Feature: DISPOFF/DISPON",
                "UserForm tip: ScrollBars = fmScrollBarsBoth",
                "Performance tip: Turn off DisplayPageBreaks"
            ]
        }
        self.wfile.write(json.dumps(response).encode())

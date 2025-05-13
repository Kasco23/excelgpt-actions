from http.server import BaseHTTPRequestHandler
import json

GEMS = {
    'charts': [
        'Chart.ExportAsFixedFormat for one‑line PDF export',
        '.SeriesCollection(i).ApplyDataLabels xlDataLabelsShowValue',
        'Old XLM DISPOFF/DISPON calls to animate charts flicker‑free'
    ],
    'userforms': [
        'MonthView control hides in MSCOMCT2.OCX—enable via Tools > Additional Controls',
        'Set ScrollBars = fmScrollBarsBoth for resizable forms',
        'WithEvents Scripting.Dictionary to broadcast custom events'
    ],
    'performance': [
        'ActiveSheet.DisplayPageBreaks = False to speed large sheets',
        'Batch range reads/writes via Variant = Range.Value2',
        'Application.Calculation = xlCalculationManual inside loops'
    ]
}

class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        length = int(self.headers['Content-Length'])
        body   = json.loads(self.rfile.read(length))
        topic  = body.get('topic', '').lower()
        max_n  = body.get('maxItems', 5)
        gems   = GEMS.get(topic, ['No gems for that topic yet!'])[:max_n]

        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps({'gems': gems}).encode())

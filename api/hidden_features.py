from http.server import BaseHTTPRequestHandler
import json

class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        content_length = int(self.headers['Content-Length'])
        body = self.rfile.read(content_length)
        try:
            data = json.loads(body)
        except json.JSONDecodeError:
            data = {}

        topic = data.get("topic", "").lower()
        max_items = data.get("maxItems", 5)

        # Hidden features database
        GEMS = {
            "charts": [
                "Use Chart.ExportAsFixedFormat for PDF export.",
                "ApplyDataLabels with xlDataLabelsShowValue to show values.",
                "Use XLM macro DISPOFF/DISPON to animate charts without flicker."
            ],
            "userforms": [
                "MonthView control is in MSCOMCT2.OCX — register it, then enable via Tools > Additional Controls.",
                "Set ScrollBars = fmScrollBarsBoth for scrollable/resizable forms.",
                "Use WithEvents with Scripting.Dictionary to trigger custom UI events."
            ],
            "performance": [
                "Disable DisplayPageBreaks during long operations for better speed.",
                "Use Application.Calculation = xlCalculationManual inside loops.",
                "Batch-read/write with Variant = Range.Value2 instead of cell-by-cell."
            ]
        }

        gems = GEMS.get(topic, ["No hidden features available for this topic."])
        gems = gems[:max_items]

        # Response
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.end_headers()
        self.wfile.write(json.dumps({"gems": gems}).encode())

    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain")
        self.end_headers()
        self.wfile.write(b"This endpoint expects a POST request with a topic (e.g. 'charts').")

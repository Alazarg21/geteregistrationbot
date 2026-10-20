
from http.server import BaseHTTPRequestHandler
import json


class handler(BaseHTTPRequestHandler):
    def do_GET(self):
        body = json.dumps({
            "status": "ok",
            "message": "Gete Registration Bot API is running"
        }).encode("utf-8")

        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        body = json.dumps({
            "error": "Telegram webhook is not configured yet"
        }).encode("utf-8")

        self.send_response(501)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

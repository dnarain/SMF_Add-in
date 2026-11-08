#!/usr/bin/env python3
"""Serve deterministic parser inputs on loopback; no market data is simulated."""
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
from pathlib import Path
import sys
import base64

HTML = b'<html><body><table><tr><td>Label</td><td>42</td></tr><tr><td>Earnings</td><td>Nov 20 AMC</td></tr></table></body></html>'
CSV = b'Date,Close,Volume\n2026-11-05,100.5,1000\n2026-11-06,102.25,1200\n'
PNG = base64.b64decode('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+ip1sAAAAASUVORK5CYII=')
class FixtureHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        route = self.path.split('?', 1)[0]
        status, kind, body = {
            '/text': (200, 'text/plain', b'SMF fixture response'),
            '/denied': (403, 'text/plain', b'fixture denied'),
            '/html': (200, 'text/html', HTML),
            '/data.csv': (200, 'text/csv', CSV),
            '/image.png': (200, 'image/png', PNG),
        }.get(route, (404, 'text/plain', b'fixture missing'))
        self.send_response(status)
        self.send_header('Content-Type', kind)
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)
    def log_message(self, *args):
        pass
if __name__ == '__main__':
    server = ThreadingHTTPServer(('127.0.0.1', 0), FixtureHandler)
    Path(sys.argv[1]).write_text(json.dumps({'base_url': 'http://127.0.0.1:' + str(server.server_port)}))
    print('Loopback fixture server ready', flush=True)
    server.serve_forever()

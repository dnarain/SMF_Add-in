#!/usr/bin/env python3
"""Serve explicit parser fixtures on loopback; never simulate a live provider."""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import json
import os
from pathlib import Path
import struct
from urllib.parse import unquote, urlsplit
import zlib


def builtin_fixtures():
    def chunk(kind, data):
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data) & 0xffffffff)
    png = (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 2, 2, 8, 2, 0, 0, 0))
           + chunk(b'IDAT', zlib.compress(b'\0' + b'\x20\x80\xc0' * 2 + b'\0' + b'\xf0\xf0\xf0' * 2)) + chunk(b'IEND', b''))
    return {'/text': (200, 'text/plain; charset=utf-8', b'SMF fixture response'),
            '/denied': (403, 'text/plain; charset=utf-8', b'fixture denied'),
            '/html': (200, 'text/html; charset=utf-8', b'<html><body><table><tr><td>Label</td><td>42</td></tr><tr><td>Earnings</td><td>Nov 20 AMC</td></tr></table></body></html>'),
            '/data.csv': (200, 'text/csv; charset=utf-8', b'Date,Close,Volume\n2026-11-05,100.5,1000\n2026-11-06,102.25,1200\n'),
            '/image.png': (200, 'image/png', png)}


def load_fixtures(root: Path | None):
    if root is None:
        return builtin_fixtures()
    root = root.resolve()
    manifest = json.loads((root / 'routes.json').read_text(encoding='utf-8'))
    fixtures = {}
    for route, entry in manifest.items():
        if not route.startswith('/') or '?' in route or '#' in route or '..' in route.split('/'):
            raise ValueError('Invalid fixture route: ' + route)
        file = (root / entry['file']).resolve()
        if root not in file.parents or not file.is_file():
            raise ValueError('Fixture route must name a file within --root: ' + route)
        status = int(entry['status'])
        if status not in (200, 403):
            raise ValueError('Only the reviewed 200/403 fixture statuses are supported')
        content_type = entry['content_type']
        if '\r' in content_type or '\n' in content_type:
            raise ValueError('Invalid fixture content type')
        fixtures[route] = (status, content_type, file.read_bytes())
    return fixtures


def handler_for(fixtures):
    class FixtureHandler(BaseHTTPRequestHandler):
        def do_GET(self):
            route = unquote(urlsplit(self.path).path)
            status, kind, body = fixtures.get(route, (404, 'text/plain; charset=utf-8', b'fixture missing'))
            self.send_response(status)
            self.send_header('Content-Type', kind)
            self.send_header('Content-Length', str(len(body)))
            self.send_header('Cache-Control', 'no-store')
            self.end_headers()
            self.wfile.write(body)

        def log_message(self, *args):
            pass
    return FixtureHandler


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('legacy_output', nargs='?', type=Path, help='Legacy explicit config-output path')
    parser.add_argument('--output', type=Path, help='Write the current loopback URL and PID to this JSON file')
    parser.add_argument('--root', type=Path, help='Prepared fixture directory containing routes.json')
    parser.add_argument('--port', type=int, default=0, help='Loopback port; zero chooses a free port')
    args = parser.parse_args()
    if bool(args.output) == bool(args.legacy_output):
        parser.error('Supply exactly one --output path or legacy positional output path')
    if not 0 <= args.port <= 65535:
        parser.error('--port must be between 0 and 65535')
    output = (args.output or args.legacy_output).resolve()
    try:
        fixtures = load_fixtures(args.root)
    except (ValueError, KeyError, OSError) as error:
        parser.error(str(error))
    try:
        server = ThreadingHTTPServer(('127.0.0.1', args.port), handler_for(fixtures))
    except OSError as error:
        parser.error('Cannot start the loopback fixture server: ' + str(error))
    state = {'schema_version': 1, 'base_url': 'http://127.0.0.1:' + str(server.server_port),
             'host': '127.0.0.1', 'port': server.server_port, 'pid': os.getpid(),
             'fixture_root': str(args.root.resolve()) if args.root else None,
             'started_utc': datetime.now(timezone.utc).isoformat(),
             'routes': sorted(fixtures), 'market_data_simulated': False}
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(output.name + '.' + str(os.getpid()) + '.tmp')
    temporary.write_text(json.dumps(state, indent=2) + '\n', encoding='utf-8')
    temporary.replace(output)
    print('Loopback fixture server ready: ' + state['base_url'], flush=True)
    try:
        server.serve_forever(poll_interval=0.2)
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == '__main__':
    main()

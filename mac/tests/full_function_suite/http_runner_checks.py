#!/usr/bin/env python3
"""Check HTTP evidence output paths and responses against loopback fixtures."""
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import hashlib
import importlib.util
import json
import subprocess
import sys
import tempfile
import threading
import unittest

RUNNER = Path(__file__).with_name('run_datasource_http.py')

class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        denied = self.path == '/denied'
        self.send_response(403 if denied else 200)
        self.send_header('Content-Type', 'text/plain')
        self.send_header('Set-Cookie', 'anonymous=fixture-value')
        self.end_headers()
        self.wfile.write(b'Access Denied' if denied else b'fixture exact value 42')
    def log_message(self, *args):
        pass

class HTTPRunnerChecks(unittest.TestCase):
    def test_import_has_no_output_directory_side_effect(self):
        spec=importlib.util.spec_from_file_location('smf_http_evidence',RUNNER)
        module=importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        self.assertIsNone(module.OUTPUT)

    def test_explicit_output_outside_checkout_and_redacted_headers(self):
        server=ThreadingHTTPServer(('127.0.0.1',0),Handler)
        worker=threading.Thread(target=server.serve_forever,daemon=True);worker.start()
        try:
            with tempfile.TemporaryDirectory(prefix='smf HTTP runner ') as temp:
                root=Path(temp); output=root/'results with spaces'
                base='http://127.0.0.1:'+str(server.server_port)
                inventory=root/'inventory.json';meta=root/'meta.json'
                inventory.write_text(json.dumps([{'id':'fixture','kind':'test_fixture','probe_urls':[base+'/ok',base+'/denied']}]))
                meta.write_text(json.dumps({'unreachable_option_references':[]}))
                subprocess.run([sys.executable,str(RUNNER),'--output',str(output),'--workers','2','--inventory',str(inventory),'--metadata',str(meta)],check=True,capture_output=True,text=True,timeout=20)
                records=json.loads((output/'results.json').read_text())
                self.assertEqual(len(records),2)
                self.assertEqual({r['response']['http_status'] for r in records},{200,403})
                for record in records:
                    response=record['response']
                    self.assertFalse(Path(response['body_path']).is_absolute())
                    payload=(output/response['body_path']).read_bytes()
                    self.assertEqual(hashlib.sha256(payload).hexdigest(),response['body_sha256'])
                    headers=(output/response['headers_path']).read_text()
                    self.assertIn('[redacted]',headers)
                    self.assertNotIn('fixture-value',headers)
                metadata=json.loads((output/'run-metadata.json').read_text())
                self.assertEqual(metadata['schema_version'],2)
                self.assertEqual(metadata['path_base'],'this_output_folder')
        finally:
            server.shutdown();server.server_close()

    def test_bad_worker_count_does_not_create_output(self):
        with tempfile.TemporaryDirectory(prefix='smf HTTP rejected ') as temp:
            output=Path(temp)/'not created'
            result=subprocess.run([sys.executable,str(RUNNER),'--output',str(output),'--workers','0'],capture_output=True,text=True)
            self.assertNotEqual(result.returncode,0)
            self.assertFalse(output.exists())

if __name__=='__main__':
    unittest.main(verbosity=2)

"""Exercise the actual AppleScript/curl helper against a local HTTP server."""
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import json
import subprocess
import threading
import tempfile
import unittest

HELPER = Path(__file__).resolve().parents[1] / "SMFMac.applescript"
COMPILED_HELPER = None


class FixtureServer(BaseHTTPRequestHandler):
    def log_message(self, *_):
        pass

    def do_GET(self):
        if self.path == "/redirect":
            self.send_response(302)
            self.send_header("Location", "/text")
            self.end_headers()
            return
        if self.path == "/redirect-file":
            self.send_response(302)
            self.send_header("Location", "file:///etc/hosts")
            self.end_headers()
            return
        status = 403 if self.path == "/denied" else 200
        body = {
            "/text": "Revenue $1,234.50\nEPS €2.35\n",
            "/empty": "",
            "/denied": "<html>Access denied</html>",
            "/large": "<td>42</td>" * 20000,
        }.get(self.path, self.path)
        self.send_response(status)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write(body.encode("utf-8"))

    def do_POST(self):
        self.send_response(200)
        self.end_headers()
        self.wfile.write(b"POST")


def run_request(request):
    result = subprocess.run(
        ["/usr/bin/osascript", str(HELPER), request],
        capture_output=True, text=True, timeout=35, check=True,
    )
    # osascript appends one newline to the returned string.
    output = result.stdout.removesuffix("\n")
    protocol, status, body = output.split("\n", 2)
    if protocol != "SMFHTTP/1":
        raise AssertionError(output[:300])
    return int(status), body


def run_handler(name, *values):
    script = (
        "set helper to load script POSIX file " + json.dumps(str(COMPILED_HELPER))
        + "\nreturn helper's " + name + "(" + ", ".join(json.dumps(value) for value in values) + ")"
    )
    result = subprocess.run(
        ["/usr/bin/osascript", "-e", script],
        capture_output=True, timeout=10, check=True,
    )
    return result.stdout.decode("utf-8").removesuffix("\n")


class TransportTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        global COMPILED_HELPER
        cls.tempdir = tempfile.TemporaryDirectory(prefix="smf-http-tests-")
        COMPILED_HELPER = Path(cls.tempdir.name) / "SMFMac.scpt"
        subprocess.run(["/usr/bin/osacompile", "-o", str(COMPILED_HELPER), str(HELPER)],
                       check=True, capture_output=True, text=True, timeout=15)
        cls.server = ThreadingHTTPServer(("127.0.0.1", 0), FixtureServer)
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        cls.base = "http://127.0.0.1:" + str(cls.server.server_port)

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.tempdir.cleanup()

    def get(self, path, method="GET"):
        return run_request(method + "\n" + self.base + path)

    def test_utf8_and_line_endings(self):
        self.assertEqual(self.get("/text"), (200, "Revenue $1,234.50\nEPS €2.35\n"))

    def test_empty_success(self):
        self.assertEqual(self.get("/empty"), (200, ""))

    def test_http_error_is_distinct(self):
        self.assertEqual(self.get("/denied"), (403, "<html>Access denied</html>"))

    def test_redirect(self):
        self.assertEqual(self.get("/redirect"), self.get("/text"))

    def test_large_page(self):
        status, body = self.get("/large")
        self.assertEqual(status, 200)
        self.assertEqual(body, "<td>42</td>" * 20000)

    def test_post(self):
        self.assertEqual(self.get("/post", "POST"), (200, "POST"))

    def test_type4_on_other_sites_uses_get(self):
        self.assertEqual(self.get("/text", "YAHOO"), self.get("/text"))

    def test_literal_quote_and_encoded_shell_characters(self):
        path = "/quote?x='&value=%24%28printf%20injected%29"
        self.assertEqual(self.get(path), (200, path))

    def test_url_newline_rejected(self):
        status, body = run_request("GET\n" + self.base + "/text\nextra")
        self.assertEqual(status, 0)
        self.assertIn("line break", body)

    def test_unsupported_protocol_rejected(self):
        self.assertEqual(run_request("GET\nfile:///etc/hosts")[0], 0)

    def test_redirect_cannot_read_local_file(self):
        self.assertEqual(self.get("/redirect-file")[0], 0)

    def test_unknown_method_rejected(self):
        self.assertEqual(self.get("/text", "DELETE")[0], 0)

    def test_invalid_request_rejected(self):
        self.assertEqual(run_request("invalid")[0], 0)

    def test_crumb_replacement(self):
        self.assertEqual(run_handler("removeCrumb", "https://query1.finance.yahoo.com/v7/quote?symbols=AAPL&crumb=old&x=1"),
                         "https://query1.finance.yahoo.com/v7/quote?symbols=AAPL&x=1")
        self.assertEqual(run_handler("removeCrumb", "https://query1.finance.yahoo.com/v7/quote?crumb="),
                         "https://query1.finance.yahoo.com/v7/quote")

    def test_crumb_encoding(self):
        self.assertEqual(run_handler("encodeQueryValue", "a/b+=?&%"), "a%2Fb%2B%3D%3F%26%25")

    def test_cookie_request_host_check(self):
        self.assertEqual(run_handler("isYahooAPI", "https://query1.finance.yahoo.com/v7/options/AAPL"), "true")
        self.assertEqual(run_handler("isYahooAPI", "https://query1.finance.yahoo.com.evil.test/"), "false")
        self.assertEqual(run_handler("isYahooAPI", "http://query1.finance.yahoo.com/v7/options/AAPL"), "false")

    def test_module_copy_preserves_bytes_and_class_attributes(self):
        project = Path(self.tempdir.name) / "fixture-project"
        (project / "code").mkdir(exist_ok=True, parents=True)
        body = (b'VERSION 1.0 CLASS\r\nAttribute VB_Name = "Dictionary"\r\n'
                b'Attribute Item.VB_UserMemId = 0\r\nAttribute NewEnum.VB_UserMemId = -4\r\n')
        (project / "code/Dictionary.cls").write_bytes(body)
        destination = Path(self.tempdir.name) / "stage-copy"
        result = run_handler("copyProjectModuleTo", str(project), "Dictionary.cls", str(destination))
        self.assertTrue(result.endswith(":Dictionary.cls"))
        self.assertEqual((destination / "Dictionary.cls").read_bytes(), body)

    def test_module_copy_rejects_traversal(self):
        with self.assertRaises(subprocess.CalledProcessError):
            run_handler("copyProjectModuleTo", self.tempdir.name, "../Dictionary.cls", self.tempdir.name)

    def test_module_copy_rejects_other_file_types(self):
        with self.assertRaises(subprocess.CalledProcessError):
            run_handler("copyProjectModuleTo", self.tempdir.name, "secret.txt", self.tempdir.name)

    def test_element_file_utf8_and_line_endings(self):
        folder = Path(self.tempdir.name) / "element-fixture"
        folder.mkdir(exist_ok=True)
        text = "4072;FinViz;Earnings Date;=test()\r\n4001;Test;€;=test()\n"
        (folder / "smf-elements-4.txt").write_bytes(text.encode("utf-8"))
        self.assertEqual(run_handler("readElementFile", str(folder) + "\n4"), text)

    def test_custom_element_suffix(self):
        folder = Path(self.tempdir.name) / "custom-element-fixture"
        folder.mkdir(exist_ok=True)
        (folder / "smf-elements-20-FinViz_v2.txt").write_text("19172;FinViz;Earnings\n")
        self.assertEqual(run_handler("readElementFile", str(folder) + "\n20-FinViz_v2"),
                         "19172;FinViz;Earnings\n")

    def test_missing_element_file_is_empty(self):
        self.assertEqual(run_handler("readElementFile", self.tempdir.name + "\n21"), "")

    def test_element_file_rejects_traversal(self):
        with self.assertRaises(subprocess.CalledProcessError):
            run_handler("readElementFile", self.tempdir.name + "\n../4")

    def test_element_location_file(self):
        folder = Path(self.tempdir.name) / "location-fixture"
        folder.mkdir(exist_ok=True)
        self.assertEqual(run_handler("readElementLocation", str(folder)), "")
        (folder / "smf-Elements-Location.txt").write_bytes(b"Local\r\n")
        self.assertEqual(run_handler("readElementLocation", str(folder)), "Local\r\n")


if __name__ == "__main__":
    unittest.main(verbosity=2)

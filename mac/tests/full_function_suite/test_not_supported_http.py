"""Check retired HTTP skips and active probes without making any HTTP request."""
from contextlib import redirect_stdout
from io import StringIO
import hashlib
import importlib.util
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest import mock

RUNNER = Path(__file__).with_name("run_datasource_http.py")
ROOT = Path(__file__).resolve().parents[3]


class NotSupportedHTTPTests(unittest.TestCase):
    def setUp(self):
        spec = importlib.util.spec_from_file_location("smf_not_supported_http_runner", RUNNER)
        self.runner = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(self.runner)
        self.temp = tempfile.TemporaryDirectory(prefix="smf retired HTTP mock ")
        self.addCleanup(self.temp.cleanup)
        self.output = Path(self.temp.name)
        self.runner.OUTPUT = self.output
        self.support = self.runner.load_support_registry(ROOT / "mac/not_supported.json")
        # Every test rejects an accidental real curl subprocess, including active-route tests.
        self.no_subprocess = mock.patch.object(self.runner.subprocess, "run", side_effect=AssertionError("Actual HTTP is forbidden in these tests"))
        self.no_subprocess.start()
        self.addCleanup(self.no_subprocess.stop)

    def job(self, url, mode="GET"):
        return {"url": url, "mode": mode, "source_ids": ["fixture"], "kind": "test_fixture"}

    def fake_curl(self, url, prefix, cookie=None):
        body = prefix.with_suffix(".body")
        header = prefix.with_suffix(".headers")
        payload = b"FAKE-CRUMB" if "/v1/test/getcrumb" in url else b"mock provider response"
        body.write_bytes(payload)
        header.write_text("HTTP/1.1 200 OK\n")
        return {"requested_url": url, "effective_url": url, "http_status": 200, "curl_exit": 0,
                "body_path": str(body.relative_to(self.output)), "headers_path": str(header.relative_to(self.output)),
                "body_sha256": hashlib.sha256(payload).hexdigest(), "bytes": len(payload),
                "set_cookie_present": True}

    def assert_skip(self, record, entry_id):
        self.assertEqual(record["support_id"], entry_id)
        self.assertEqual(record["support_status"], "Not Supported")
        self.assertIs(record["request_performed"], False)
        self.assertEqual(record["network_requests"], 0)
        response = record["response"]
        self.assertEqual(response["http_status"], 0)
        self.assertEqual(response["support_id"], entry_id)
        self.assertEqual(response["support_status"], "Not Supported")
        self.assertIs(response["request_performed"], False)
        self.assertEqual(response["network_requests"], 0)
        self.assertEqual(response["body_path"], "")
        self.assertEqual(response["headers_path"], "")
        self.assertEqual(response["body_sha256"], "")
        self.assertIsNone(response["curl_exit"])
        self.assertNotIn("auth", record)

    def test_every_registered_route_skips_without_curl_or_response_files(self):
        with mock.patch.object(self.runner, "curl", side_effect=AssertionError("Retired route called curl")) as curl:
            for index, (entry_id, host, path) in enumerate(self.support["url_rules"], 1):
                with self.subTest(entry_id=entry_id, host=host, path=path):
                    record = self.runner.run_job(index, self.job("https://" + host + path + "?symbol=AAPL#ignored"), self.support)
                    self.assert_skip(record, entry_id)
                    self.assertEqual(json.loads((self.output / f"{index:03d}.json").read_text()), record)
            curl.assert_not_called()
        self.assertFalse(list(self.output.glob("*.body")))
        self.assertFalse(list(self.output.glob("*.headers")))

    def test_registered_yahoo_mode_skips_before_cookie_handshake(self):
        url = "https://query1.finance.yahoo.com/v8/finance/chart/AAPL"
        support = {**self.support, "entries": {**self.support["entries"], "fixture-retired": {"reason": "Synthetic registered test route"}},
                   "url_rules": (("fixture-retired", "query1.finance.yahoo.com", "/v8/finance/chart/aapl"),)}
        with mock.patch.object(self.runner, "curl", side_effect=AssertionError("Skip entered Yahoo authentication")) as curl, \
                mock.patch.object(self.runner.tempfile, "TemporaryDirectory", side_effect=AssertionError("Skip created Yahoo cookie directory")):
            record = self.runner.run_job(1, self.job(url, "YAHOO"), support)
            self.assert_skip(record, "fixture-retired")
            curl.assert_not_called()

    def test_userinfo_and_one_trailing_dns_dot_follow_vba_matching(self):
        url = "HTTPS://fixture:password@QUERY.YAHOOAPIS.COM.:443/v1/public/yql?query=ignored"
        with mock.patch.object(self.runner, "curl", side_effect=AssertionError("Registered userinfo URL called curl")):
            self.assert_skip(self.runner.run_job(1, self.job(url), self.support), "yahoo-yql")

    def test_modern_yahoo_history_remains_an_active_get_probe(self):
        url = "https://query1.finance.yahoo.com/v8/finance/chart/AAPL?interval=1d"
        with mock.patch.object(self.runner, "curl", side_effect=self.fake_curl) as curl:
            record = self.runner.run_job(1, self.job(url), self.support)
        self.assertEqual(curl.call_count, 1)
        self.assertEqual(record["response"]["http_status"], 200)
        self.assertNotIn("support_id", record)
        self.assertNotIn("auth", record)

    def test_modern_yahoo_history_keeps_existing_authenticated_probe(self):
        url = "https://query1.finance.yahoo.com/v8/finance/chart/AAPL?interval=1d"
        with mock.patch.object(self.runner, "curl", side_effect=self.fake_curl) as curl:
            record = self.runner.run_job(1, self.job(url, "YAHOO"), self.support)
        self.assertEqual(curl.call_count, 3)
        self.assertEqual(curl.call_args_list[0].args[0], "https://fc.yahoo.com")
        self.assertEqual(curl.call_args_list[1].args[0], "https://query1.finance.yahoo.com/v1/test/getcrumb")
        self.assertTrue(record["auth"]["crumb_valid"])
        self.assertEqual(record["response"]["requested_url"], url)
        self.assertIn("crumb=[redacted]", record["response"]["effective_url"])
        self.assertEqual((self.output / "001-crumb.body").read_text(), "[crumb redacted]\n")
        self.assertNotIn("support_id", record)

    def test_unregistered_host_path_and_current_google_neighbors_are_active(self):
        urls = ["https://download.finance.yahoo.com.example.invalid/d/quotes.csv",
                "https://download.finance.yahoo.com/d/quotes.csv-extra",
                "https://query.yahooapis.com/v1/public/yql-extra",
                "https://query.yahooapis.com@example.invalid/v1/public/yql",
                "https://www.google.com/finance/quote/AAPL?next=https://query.yahooapis.com/v1/public/yql"]
        with mock.patch.object(self.runner, "curl", side_effect=self.fake_curl) as curl:
            for index, url in enumerate(urls, 1):
                with self.subTest(url=url):
                    record = self.runner.run_job(index, self.job(url), self.support)
                    self.assertEqual(record["response"]["http_status"], 200)
                    self.assertNotIn("support_id", record)
            self.assertEqual(curl.call_count, len(urls))

    def test_http_denial_is_not_inferred_as_retirement(self):
        def denied(url, prefix, cookie=None):
            response = self.fake_curl(url, prefix, cookie)
            response["http_status"] = 403
            return response
        url = "https://query1.finance.yahoo.com/v7/finance/quote?symbols=AAPL"
        with mock.patch.object(self.runner, "curl", side_effect=denied) as curl:
            record = self.runner.run_job(1, self.job(url), self.support)
        self.assertEqual(curl.call_count, 1)
        self.assertEqual(record["response"]["http_status"], 403)
        self.assertNotIn("support_id", record)
        self.assertNotIn("support_status", record["response"])

    def test_cli_explicit_registry_hash_and_skip_record_are_saved(self):
        registry = self.output / "explicit-policy.json"
        source = json.loads((ROOT / "mac/not_supported.json").read_text())
        source["reviewed"] = "fixture review marker"
        registry.write_text(json.dumps(source))
        inventory = self.output / "inventory.json"
        inventory.write_text(json.dumps([{"id": "fixture", "kind": "retired_test", "probe_urls": ["https://query.yahooapis.com/v1/public/yql?q=test"], "mode": "YAHOO"}]))
        metadata = self.output / "metadata.json"
        metadata.write_text(json.dumps({"unreachable_option_references": []}))
        output = self.output / "records"
        argv = [str(RUNNER), "--output", str(output), "--workers", "1", "--inventory", str(inventory),
                "--metadata", str(metadata), "--support-registry", str(registry)]
        with mock.patch.object(sys, "argv", argv), redirect_stdout(StringIO()), \
                mock.patch.object(self.runner, "curl", side_effect=AssertionError("CLI skip attempted curl")) as curl:
            self.runner.main()
            curl.assert_not_called()
        manifest = json.loads((output / "run-metadata.json").read_text())
        self.assertEqual(manifest["support_registry_sha256"], hashlib.sha256(registry.read_bytes()).hexdigest())
        self.assertEqual(manifest["schema_version"], 2)
        records = json.loads((output / "results.json").read_text())
        self.assertEqual(len(records), 1)
        self.assert_skip(records[0], "yahoo-yql")

    def test_cli_rejects_unknown_registry_rule_before_any_request_or_output(self):
        registry = self.output / "invalid-policy.json"
        registry.write_text(json.dumps({"schema_version": 1, "status": "Not Supported", "entries": [],
                                        "url_rules": [{"id": "missing", "host": "example.invalid", "path": "/old"}]}))
        output = self.output / "not-created"
        with mock.patch.object(sys, "argv", [str(RUNNER), "--output", str(output), "--support-registry", str(registry)]), \
                mock.patch.object(self.runner, "curl", side_effect=AssertionError("Invalid registry attempted curl")):
            with self.assertRaisesRegex(ValueError, "unknown entry ID"):
                self.runner.main()
        self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main(verbosity=2)

#!/usr/bin/env python3
"""Check fixture paths produced by the actual suite generator; Excel is not run."""
from pathlib import Path
import json
import re
import subprocess
import sys
import tempfile
import unittest

BUILDER = Path(__file__).with_name("build_suite.py")


class GeneratorPathChecks(unittest.TestCase):
    def check_suite(self, suite_argument, working_directory):
        folder = Path(working_directory)
        suite = (folder / suite_argument).resolve()
        suite.mkdir(parents=True)
        cases = [{
            "id": "Generator_AbsoluteFixturePath",
            "function": "RCHCreateComment",
            "group": "generator_fixture",
            "description": "Retain an absolute fixture filename in generated VBA.",
            "declarations": ["Dim fixtureFile As String"],
            "setup": ['fixtureFile = "{{SUITE_DIR}}/fixture.png"'],
            "expression": "fixtureFile",
            "object_result": False,
            "check": "Len(CStr(result)) > 0",
            "expected_text": "An absolute filename inside the selected suite folder.",
            "cleanup": [],
        }]
        case_file = folder / "case-input.json"
        case_file.write_text(json.dumps(cases))
        server_config = folder / "server.json"
        server_config.write_text(json.dumps({"base_url": "http://127.0.0.1:1"}))
        subprocess.run([sys.executable, str(BUILDER), "--suite", str(suite_argument),
                        "--server-config", str(server_config), "--case-file", str(case_file)],
                       cwd=folder, capture_output=True, text=True, check=True, timeout=20)
        generated = json.loads((suite / "test-cases.json").read_text())
        assignment = generated[0]["setup"][0]
        match = re.fullmatch(r'fixtureFile = "([^"]+)"', assignment)
        self.assertIsNotNone(match)
        self.assertTrue(Path(match[1]).is_absolute())
        self.assertEqual(Path(match[1]), suite / "fixture.png")
        self.assertIn(assignment, (suite / "modSMFFunctionSuite.bas").read_text())
        self.assertNotIn("{{SUITE_DIR}}", (suite / "modSMFFunctionSuite.bas").read_text())

    def test_relative_suite_argument_produces_absolute_fixture_filename(self):
        with tempfile.TemporaryDirectory(prefix="smf-generator-path-") as folder:
            self.check_suite(Path("relative suite with spaces"), folder)

    def test_absolute_suite_argument_retains_absolute_fixture_filename(self):
        with tempfile.TemporaryDirectory(prefix="smf-generator-path-") as folder:
            self.check_suite(Path(folder) / "absolute suite", folder)


if __name__ == "__main__":
    unittest.main(verbosity=2)

#!/usr/bin/env python3
"""Exercise the actual generator in fresh directories; Excel/network are not run."""
from pathlib import Path
import hashlib
import json
import re
import subprocess
import sys
import tempfile
import unittest

BUILDER = Path(__file__).with_name("build_suite.py")


class GeneratorPathChecks(unittest.TestCase):
    def case(self, case_id="Generator_AbsoluteFixturePath"):
        return {
            "id": case_id, "function": "RCHCreateComment", "group": "generator_fixture",
            "description": "Retain an absolute fixture filename in generated VBA.",
            "declarations": ["Dim fixtureFile As String"],
            "setup": ['fixtureFile = "{{SUITE_DIR}}/fixture.png"'],
            "expression": "fixtureFile", "object_result": False,
            "check": "Len(CStr(result)) > 0",
            "expected_text": "An absolute filename inside the selected suite folder.",
            "cleanup": [],
        }

    def invoke(self, folder, arguments, success=True):
        result = subprocess.run([sys.executable, str(BUILDER), *map(str, arguments)],
                                cwd=folder, capture_output=True, text=True, timeout=30)
        if success:
            self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        else:
            self.assertNotEqual(result.returncode, 0)
        return result

    def custom_inputs(self, folder, case=None):
        folder = Path(folder)
        case_file = folder / "case-input.json"
        case_file.write_text(json.dumps([case or self.case()]))
        config = folder / "server.json"
        config.write_text(json.dumps({"base_url": "http://127.0.0.1:1"}))
        return case_file, config

    def check_suite(self, suite_argument, working_directory):
        folder = Path(working_directory)
        suite = (folder / suite_argument).resolve()
        self.assertFalse(suite.exists())
        case_file, config = self.custom_inputs(folder)
        self.invoke(folder, ["--suite", suite_argument, "--server-config", config, "--case-file", case_file])
        generated = json.loads((suite / "test-cases.json").read_text())
        assignment = generated[0]["setup"][0]
        match = re.fullmatch(r'fixtureFile = "([^"]+)"', assignment)
        self.assertIsNotNone(match)
        self.assertTrue(Path(match[1]).is_absolute())
        self.assertEqual(Path(match[1]), suite / "fixture.png")
        source = (suite / "modSMFFunctionSuite.bas").read_text()
        self.assertIn(assignment, source)
        self.assertNotIn("{{SUITE_DIR}}", source)
        manifest = json.loads((suite / "run-manifest.json").read_text())
        self.assertEqual(manifest["native_execution"], "not_run")
        self.assertFalse(manifest["full_function_coverage_required"])
        self.assertEqual(manifest["case_specification_sha256"],
                         hashlib.sha256((suite / "test-cases.json").read_bytes()).hexdigest())

    def test_relative_suite_argument_creates_absolute_fixture_paths(self):
        with tempfile.TemporaryDirectory(prefix="smf-generator-path-") as folder:
            self.check_suite(Path("relative suite with spaces"), folder)

    def test_absolute_suite_argument_retains_absolute_fixture_paths(self):
        with tempfile.TemporaryDirectory(prefix="smf-generator-path-") as folder:
            self.check_suite(Path(folder) / "absolute suite", folder)

    def test_explicit_local_base_needs_no_config_or_ignored_state(self):
        with tempfile.TemporaryDirectory(prefix="smf-generator-direct-") as folder:
            case_file, config = self.custom_inputs(folder)
            config.unlink()
            suite = Path(folder) / "direct"
            self.invoke(folder, ["--suite", suite, "--local-http-base", "http://127.0.0.1:8123/",
                                 "--case-file", case_file, "--run-id", "reviewed-run-1"])
            manifest = json.loads((suite / "run-manifest.json").read_text())
            self.assertEqual(manifest["local_http_base"], "http://127.0.0.1:8123")
            self.assertEqual(manifest["run_id"], "reviewed-run-1")
            self.assertIsNone(manifest["server_config"])

    def test_missing_suite_is_rejected_instead_of_reading_old_run_state(self):
        with tempfile.TemporaryDirectory(prefix="smf-generator-required-") as folder:
            case_file, _ = self.custom_inputs(folder)
            old = Path(folder) / "mac/diagnostics/out"
            old.mkdir(parents=True)
            (old / "exhaustive-current.json").write_text('{"suite_dir": "/must-not-be-used"}')
            result = self.invoke(folder, ["--local-http-base", "http://127.0.0.1:1",
                                          "--case-file", case_file], False)
            self.assertIn("--suite", result.stderr)

    def test_endpoint_is_explicit_and_mutually_exclusive(self):
        with tempfile.TemporaryDirectory(prefix="smf-generator-endpoint-") as folder:
            case_file, config = self.custom_inputs(folder)
            common = ["--suite", "suite", "--case-file", case_file]
            self.invoke(folder, common, False)
            self.invoke(folder, common + ["--server-config", config,
                                         "--local-http-base", "http://127.0.0.1:1"], False)

    def test_nonloopback_invalid_and_credentialed_fixture_endpoints_are_rejected(self):
        with tempfile.TemporaryDirectory(prefix="smf-generator-loopback-") as folder:
            case_file, _ = self.custom_inputs(folder)
            for endpoint in ("https://example.com", "file:///tmp/example",
                             "http://name@localhost:1", "http://127.0.0.1:99999"):
                with self.subTest(endpoint=endpoint):
                    self.invoke(folder, ["--suite", "suite", "--case-file", case_file,
                                         "--local-http-base", endpoint], False)

    def test_duplicate_case_ids_are_rejected(self):
        with tempfile.TemporaryDirectory(prefix="smf-generator-duplicate-") as folder:
            case_file, _ = self.custom_inputs(folder)
            case_file.write_text(json.dumps([self.case(), self.case()]))
            result = self.invoke(folder, ["--suite", "suite", "--case-file", case_file,
                                         "--local-http-base", "http://127.0.0.1:1"], False)
            self.assertIn("Duplicate case IDs", result.stderr)

    def test_unsafe_reviewed_assertion_stays_verbatim_and_has_separate_check_phase(self):
        with tempfile.TemporaryDirectory(prefix="smf-generator-check-phase-") as folder:
            case = self.case("UnsafeCheckDiagnostic")
            case["check"] = "IsNumeric(result) And CDbl(result) = 42"
            case_file, _ = self.custom_inputs(folder, case)
            suite = Path(folder) / "suite"
            self.invoke(folder, ["--suite", suite, "--case-file", case_file,
                                 "--local-http-base", "http://127.0.0.1:1"])
            source = (suite / "modSMFFunctionSuite.bas").read_text()
            self.assertIn('phase = "CHECK"', source)
            self.assertIn("ok = " + case["check"], source)
            self.assertIn('"ASSERTION_ERROR"', source)
            self.assertIn('"CLEANUP_ERROR"', source)
            manifest = json.loads((suite / "run-manifest.json").read_text())
            self.assertEqual(manifest["assertions_with_nonshortcircuit_conversions"], [case["id"]])

    def test_fresh_default_generation_has_239_cases_121_functions_and_nine_worksheet_cases(self):
        with tempfile.TemporaryDirectory(prefix="smf-generator-default-") as folder:
            suite = Path(folder) / "fresh" / "nested suite"
            self.invoke(folder, ["--suite", suite, "--local-http-base", "http://127.0.0.1:1"])
            manifest = json.loads((suite / "run-manifest.json").read_text())
            self.assertEqual(manifest["case_count"], 239)
            self.assertEqual(manifest["public_function_count"], 121)
            self.assertEqual(manifest["covered_public_function_count"], 121)
            self.assertEqual(manifest["worksheet_history_case_count"], 9)
            self.assertTrue(manifest["full_function_coverage_required"])
            source = (suite / "modSMFFunctionSuite.bas").read_text()
            self.assertIn("Public Sub SMFRunFullSuite()", source)
            self.assertIn("Public Sub SMFTestStart(Optional", source)
            self.assertIn("3.0.2024.07.30-mac.0.5-dev", source)
            self.assertIn('control.Range("B3").Value2 = "BLOCKED_VERSION"', source)
            self.assertIn("ignoredHTTP = smfMacFetch(httpMarker)", source)
            self.assertIn('"HTTP reset marker", "HTTP case ID", "HTTP run ID", "HTTP observation"', source)
            running = source.index('ws.Cells(row, 6).Value2 = "RUNNING"')
            saved = source.index("ThisWorkbook.Save", running)
            invoke = source.index("Select Case n", running)
            self.assertLess(saved, invoke)
            resume = source.split("Public Sub SMFTestResume()", 1)[1].split("End Sub", 1)[0]
            self.assertLess(resume.index('"BLOCKED_INTERRUPTED"'), resume.index("ScheduleNext"))
            self.assertIn("no automatic retry", source)
            self.assertEqual(source.count("historyCell.Formula2 = historyFormula"), 9)
            self.assertEqual(source.count("historyCell.Calculate"), 9)
            self.assertLessEqual(max(map(len, source.splitlines())), 1023)

    def test_double_quote_output_paths_are_rejected_before_vba_is_generated(self):
        with tempfile.TemporaryDirectory(prefix="smf-generator-quote-") as folder:
            case_file, _ = self.custom_inputs(folder)
            result = self.invoke(folder, ["--suite", 'bad"literal', "--case-file", case_file,
                                         "--local-http-base", "http://127.0.0.1:1"], False)
            self.assertIn("double quotes", result.stderr)


    def test_actual_picture_cleanup_skips_absent_names_without_clearing_errors(self):
        source_cases = Path(__file__).with_name("mac_picture_cases.json")
        cases = json.loads(source_cases.read_text())
        self.assertEqual(len(cases), 10)
        with tempfile.TemporaryDirectory(prefix="smf-generator-picture-cleanup-") as folder:
            suite = Path(folder) / "suite"
            self.invoke(folder, ["--suite", suite, "--case-file", source_cases,
                                 "--local-http-base", "http://127.0.0.1:1"])
            generated = (suite / "modSMFFunctionSuite.bas").read_text()
            for index, case in enumerate(cases, 1):
                with self.subTest(case=case["id"]):
                    cleanup = case["cleanup"]
                    self.assertEqual(cleanup[0], "For Each cleanupPicture In testSheet.Shapes")
                    self.assertEqual(cleanup[1], 'If cleanupPicture.Name = "SMFCommentPicture_20_10" Then')
                    self.assertIn("SMFCommentPicture/1", cleanup[2])
                    self.assertEqual(cleanup[3:6], ["Exit For", "End If", "Next cleanupPicture"])
                    self.assertFalse(any("On Error" in line or "Err.Clear" in line for line in cleanup))
                    self.assertNotIn('Set cleanupPicture = testSheet.Shapes("SMFCommentPicture_20_10")', cleanup)
                    block = generated.split(f"Private Sub SuiteCase{index:03d}(", 1)[1].split("End Sub", 1)[0]
                    loop = block.index("For Each cleanupPicture In testSheet.Shapes")
                    next_picture = block.index("Next cleanupPicture", loop)
                    capture = block.index("cleanupNumber = Err.Number", next_picture)
                    reraised = block.index('If cleanupNumber <> 0 Then Err.Raise', capture)
                    self.assertLess(loop, next_picture)
                    self.assertLess(next_picture, capture)
                    self.assertLess(capture, reraised)
                    self.assertNotIn("Err.Clear", block[loop:capture])

    def test_explicit_cleanup_failure_is_preserved_and_reraised_as_cleanup(self):
        with tempfile.TemporaryDirectory(prefix="smf-generator-cleanup-error-") as folder:
            case = self.case("DeliberateCleanupFailure")
            case["cleanup"] = ['Err.Raise 5, , "cleanup sentinel"']
            case_file, _ = self.custom_inputs(folder, case)
            suite = Path(folder) / "suite"
            self.invoke(folder, ["--suite", suite, "--case-file", case_file,
                                 "--local-http-base", "http://127.0.0.1:1"])
            generated = (suite / "modSMFFunctionSuite.bas").read_text()
            block = generated.split("Private Sub SuiteCase001(", 1)[1].split("End Sub", 1)[0]
            failure = block.index('Err.Raise 5, , "cleanup sentinel"')
            capture = block.index("cleanupNumber = Err.Number", failure)
            reraised = block.index('If cleanupNumber <> 0 Then Err.Raise', capture)
            self.assertNotIn("Err.Clear", block[failure:capture])
            self.assertIn('" [CLEANUP]"', block[reraised:])
            self.assertIn('ws.Cells(row, 6).Value2 = "CLEANUP_ERROR"', generated)



    def test_strict_version_gate_uses_the_source_prefix_and_an_exact_token(self):
        with tempfile.TemporaryDirectory(prefix="smf-generator-version-token-") as folder:
            case_file, _ = self.custom_inputs(folder)
            suite = Path(folder) / "suite"
            self.invoke(folder, ["--suite", suite, "--case-file", case_file,
                                 "--local-http-base", "http://127.0.0.1:1"])
            generated = (suite / "modSMFFunctionSuite.bas").read_text()
            helper = generated.split("Private Function SuiteVersionToken(", 1)[1].split("End Function", 1)[0]
            prefix = re.search(r'Const prefix As String = "([^"]+)"', helper)[1]
            source = (BUILDER.resolve().parents[3] / "code/modGetElementNumber.bas").read_text()
            version_prefix = re.search(r'RCHGetElementNumber\s*=\s*"([^"]+)"\s*&\s*kVersion', source)[1]
            self.assertEqual(prefix, version_prefix)
            self.assertEqual(prefix, "Stock Market Functions add-in, Version ")
            self.assertIn('If Left$(pVersionText, Len(prefix)) <> prefix Then Exit Function', helper)
            self.assertIn('pathAt = InStr(1, pVersionText, " (", vbBinaryCompare)', helper)
            self.assertIn('If pathAt <= Len(prefix) + 1 Then Exit Function', helper)
            self.assertIn('pathAt - Len(prefix) - 1', helper)
            gate = generated.split("Private Function CheckSuiteVersion()", 1)[1].split("End Function", 1)[0]
            self.assertIn('If SuiteVersionToken(CStr(actual)) <> ExpectedSMFVersion Then', gate)
            self.assertNotIn('InStr(1, CStr(actual), ExpectedSMFVersion', gate)

    def test_generated_native_smoke_assertions_reject_version_suffix_and_path_matches(self):
        with tempfile.TemporaryDirectory(prefix="smf-generator-version-boundary-") as folder:
            version = "3.0.2024.07.30-mac.0.5-dev"
            prefix = "Stock Market Functions add-in, Version "
            stale = prefix + "3.0.2024.07.30-mac.0.4-dev (/tmp/" + version + "; test)"
            suffixed = prefix + version + "-old (/tmp; test)"
            smoke = json.loads(Path(__file__).with_name("version_gate_cases.json").read_text())
            self.assertEqual(len(smoke), 1)
            case = smoke[0]
            self.assertEqual(case["expression"], 'RCHGetElementNumber("Version")')
            self.assertIs(case["network"], False)
            self.assertEqual(case["check"], "versionGateOK")
            case_file, _ = self.custom_inputs(folder, case)
            suite = Path(folder) / "suite"
            self.invoke(folder, ["--suite", suite, "--case-file", case_file,
                                 "--local-http-base", "http://127.0.0.1:1"])
            generated = (suite / "modSMFFunctionSuite.bas").read_text()
            saved = json.loads((suite / "test-cases.json").read_text())[0]
            self.assertEqual(saved["check_setup"], case["check_setup"])
            self.assertIn(suffixed, generated)
            self.assertIn(stale, generated)
            for assertion in case["check_setup"]:
                self.assertIn(assertion, generated)
            self.assertIn('result = RCHGetElementNumber("Version")', generated)
            self.assertIn("ok = versionGateOK", generated)



if __name__ == "__main__":
    unittest.main(verbosity=2)

"""Check retirement-data generation in temporary repositories; no Excel or HTTP."""
from __future__ import annotations
import copy
import csv
import importlib.util
import io
import json
from pathlib import Path
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[2]
SPEC = importlib.util.spec_from_file_location("smf_build_not_supported_tests", ROOT / "mac/build_not_supported.py")
BUILDER = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(BUILDER)


class BuildNotSupportedTests(unittest.TestCase):
    PREFIX = ('Attribute VB_Name = "modGetElementNumber"\nOption Explicit\n'
              'Private Const ExistingValue As Long = 42\n'
              'Public Function ExistingFunction() As Long\n    ExistingFunction = ExistingValue\nEnd Function\n'
              '#If Mac Then\n')
    SUFFIX = ('\n\nPrivate Sub ExistingRuntime()\n    Debug.Print "Keep runtime logic unchanged"\nEnd Sub\n#End If\n')

    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="smf registry generator ")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        for folder in ("mac", "code", "docs/Elements"):
            (self.root / folder).mkdir(parents=True)
        self.registry = {
            "schema_version": 1, "reviewed": "2026-11-09", "platform": "Mac",
            "status": "Not Supported", "excel_error": "#N/A",
            "entries": [{"id": "fixture-retired", "functions": ["ExistingFunction"],
                         "mode": "Fixture only", "provider": "Fixture provider",
                         "reason": "Fixture source retired.", "replacement": "Use a supported fixture.",
                         "evidence": ["fixture evidence"]}],
            "url_rules": [{"id": "fixture-retired", "host": "legacy.example.invalid", "path": "/old"}],
            "numbered_fields": {"id": "obsolete-definition", "unsupported_literals": ["Obsolete", "Obsolete -- Web site no longer exists"]},
        }
        self.registry_path = self.root / "mac/not_supported.json"
        self.code_path = self.root / "code/modGetElementNumber.bas"
        self.csv_path = self.root / "mac/NotSupported.csv"
        self.write_registry()
        self.code_path.write_text(self.PREFIX + BUILDER.BEGIN + "\n' stale generated data\n" + BUILDER.END + self.SUFFIX)
        self.definition = self.root / "docs/Elements/smf-elements-2.txt"
        self.definition.write_text('625;OldProvider;Old field;="Obsolete"\n'
                                   '626;OldProvider;Gone field;="Obsolete -- Web site no longer exists"\n'
                                   '627;OldProvider;Uncertain field;="Obsolete -- No longer available?"\n')
        BUILDER.run(self.root, write=True)

    def write_registry(self, registry=None):
        self.registry_path.write_text(json.dumps(self.registry if registry is None else registry, indent=2) + "\n")

    def outputs(self):
        return self.code_path.read_bytes(), self.csv_path.read_bytes()

    def assert_invalid_preserves_outputs(self, pattern):
        before = self.outputs()
        with self.assertRaisesRegex(ValueError, pattern):
            BUILDER.run(self.root, write=True)
        self.assertEqual(self.outputs(), before)

    def test_clean_check_reports_counts_and_does_not_write(self):
        before = {p.relative_to(self.root): p.read_bytes() for p in self.root.rglob("*") if p.is_file()}
        result = BUILDER.run(self.root)
        self.assertEqual(result["result"], "PASS")
        self.assertEqual(result["mode"], "check")
        self.assertEqual(result["source_groups"], 1)
        self.assertEqual(result["exact_url_rules"], 1)
        self.assertEqual(result["default_fields"], 2)
        self.assertEqual(result["status_rows"], 3)
        self.assertEqual({p.relative_to(self.root): p.read_bytes() for p in self.root.rglob("*") if p.is_file()}, before)

    def test_tampered_csv_is_rejected_without_overwriting_it(self):
        self.csv_path.write_text(self.csv_path.read_text().replace("Fixture source retired.", "Unreviewed reason."))
        before = self.outputs()
        with self.assertRaisesRegex(ValueError, "mac/NotSupported.csv"):
            BUILDER.run(self.root)
        self.assertEqual(self.outputs(), before)

    def test_tampered_embedded_data_is_rejected_even_with_original_hash_comment(self):
        self.code_path.write_text(self.code_path.read_text().replace('"Fixture source retired."', '"Unreviewed replacement."'))
        before = self.outputs()
        with self.assertRaisesRegex(ValueError, "code/modGetElementNumber.bas"):
            BUILDER.run(self.root)
        self.assertEqual(self.outputs(), before)

    def test_duplicate_entry_ids_are_rejected_before_writing(self):
        registry = copy.deepcopy(self.registry)
        registry["entries"].append(copy.deepcopy(registry["entries"][0]))
        self.write_registry(registry)
        self.assert_invalid_preserves_outputs("Duplicate")

    def test_unknown_route_id_is_rejected_before_writing(self):
        registry = copy.deepcopy(self.registry)
        registry["url_rules"][0]["id"] = "missing-entry"
        self.write_registry(registry)
        self.assert_invalid_preserves_outputs("unknown")

    def test_invalid_entry_ids_are_rejected_before_writing(self):
        for invalid in ("", "   ", None, 17, "bad\nid", "non-ascii-\u2603"):
            with self.subTest(entry_id=invalid):
                registry = copy.deepcopy(self.registry)
                registry["entries"][0]["id"] = invalid
                registry["url_rules"][0]["id"] = invalid
                self.write_registry(registry)
                self.assert_invalid_preserves_outputs("nonempty|one-line ASCII")

    def test_reversed_markers_are_rejected_before_writing(self):
        self.code_path.write_text(self.PREFIX + BUILDER.END + "\n' reversed data markers\n" + BUILDER.BEGIN + self.SUFFIX)
        self.assert_invalid_preserves_outputs("order")

    def test_missing_or_repeated_markers_are_rejected_before_writing(self):
        original = self.code_path.read_text()
        for altered in (original.replace(BUILDER.BEGIN, "' missing marker"), original + BUILDER.END):
            with self.subTest(code=altered[-60:]):
                self.code_path.write_text(altered)
                self.assert_invalid_preserves_outputs("markers")

    def test_custom_definition_replaces_retired_default_and_alternative_file_is_ignored(self):
        custom = self.root / "docs/Elements/smf-elements-20.txt"
        custom.write_text('625;ReplacementProvider;Working replacement;=1+2\n')
        (self.root / "docs/Elements/smf-elements-20-FinViz_v2.txt").write_text('625;AlternativeProvider;Not the selected file;="Obsolete"\n')
        result = BUILDER.run(self.root, write=True)
        rows = list(csv.DictReader(io.StringIO(self.csv_path.read_text())))
        fields = {row["ID"]: row for row in rows if row["ID"].startswith("field-")}
        self.assertNotIn("field-625", fields)
        self.assertNotIn("field-627", fields)
        self.assertEqual(set(fields), {"field-626"})
        self.assertEqual(fields["field-626"]["Evidence"], "docs/Elements/smf-elements-2.txt:2")
        self.assertEqual(result["default_fields"], 1)
        custom.write_text('625;ReplacementProvider;New explicit retirement;="Obsolete"\n')
        BUILDER.run(self.root, write=True)
        fields = {row["ID"]: row for row in csv.DictReader(io.StringIO(self.csv_path.read_text())) if row["ID"].startswith("field-")}
        self.assertEqual(fields["field-625"]["Provider"], "ReplacementProvider")
        self.assertIn("New explicit retirement", fields["field-625"]["Function or field"])
        self.assertEqual(fields["field-625"]["Evidence"], "docs/Elements/smf-elements-20.txt:1")

    def test_write_changes_only_generated_region_and_csv(self):
        registry = copy.deepcopy(self.registry)
        registry["entries"][0]["reason"] = 'The source says "retired"; use another source.'
        self.write_registry(registry)
        registry_before = self.registry_path.read_bytes()
        definition_before = self.definition.read_bytes()
        BUILDER.run(self.root, write=True)
        code = self.code_path.read_text()
        self.assertEqual(code[:code.index(BUILDER.BEGIN)], self.PREFIX)
        self.assertEqual(code[code.index(BUILDER.END) + len(BUILDER.END):], self.SUFFIX)
        self.assertEqual(code.count(BUILDER.BEGIN), 1)
        self.assertEqual(code.count(BUILDER.END), 1)
        self.assertIn('""retired""', code)
        self.assertEqual(self.registry_path.read_bytes(), registry_before)
        self.assertEqual(self.definition.read_bytes(), definition_before)
        self.assertEqual(BUILDER.run(self.root)["result"], "PASS")

    def test_invalid_url_rules_do_not_change_outputs(self):
        for field, value in (("host", "UPPER.example.invalid"), ("path", "relative"),
                             ("path", "/old?query=value"), ("path", "/old#fragment")):
            with self.subTest(field=field, value=value):
                registry = copy.deepcopy(self.registry)
                registry["url_rules"][0][field] = value
                self.write_registry(registry)
                self.assert_invalid_preserves_outputs("exact path")


if __name__ == "__main__":
    unittest.main(verbosity=2)

#!/usr/bin/env python3
"""Report one explicitly selected native suite/workbook. No network or Excel calls.

Last-HTTP columns are diagnostics unless N:Q bind a real request to this case
and run. A successful endpoint-content check is not a financial extraction.
"""
from __future__ import annotations

import argparse
from collections import Counter
import csv
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import re
import sys
from urllib.parse import urlsplit

from read_results import WorkbookReadError, read, redact_anonymous, write_private_json

FAILURES = {"CHECK_FAILED", "RUNTIME_ERROR", "SETUP_ERROR", "ASSERTION_ERROR",
            "CALL_ERROR", "CHECK_ERROR", "HARNESS_ERROR", "CLEANUP_ERROR"}
UNFINISHED = {"", "MISSING", "PENDING", "RUNNING"}
UNAVAILABLE_HTTP = {401, 403, 404, 407, 408, 410, 429, 451}
BLOCKING = ("access denied", "just a moment...", "cf-chl-", "captcha", "unauthorized")
VERSION_PREFIX = "Stock Market Functions add-in, Version "


def load(path):
    return json.loads(Path(path).read_text(encoding="utf-8"))


def metadata_value(metadata, *keys):
    normalized = {re.sub(r"[^a-z0-9]", "", str(k).lower()): v for k, v in metadata.items()}
    for key in keys:
        value = normalized.get(re.sub(r"[^a-z0-9]", "", key.lower()))
        if value not in (None, ""):
            return str(value)
    return ""


def observed_version_token(actual):
    """Match the native gate's fixed prefix, delimiter and bare full-version token."""
    if not isinstance(actual, str) or not actual.startswith(VERSION_PREFIX):
        return ""
    delimiter = actual.find(" (", len(VERSION_PREFIX))
    if delimiter < 0:
        return ""
    return actual[len(VERSION_PREFIX):delimiter]


def version_matches(expected, actual):
    return bool(expected and observed_version_token(actual) == expected)


def coverage_kind(case):
    mode = case.get("test_mode", "")
    if mode == "live_endpoint_request" or case.get("group") == "live_datasources":
        return "HTTP_ENDPOINT"
    if case.get("fixture_source") or case.get("domain") == "loopback":
        return "FIXTURE"
    if mode.startswith("worksheet") or any(".Formula2" in step for step in case.get("call_steps", [])):
        return "WORKSHEET"
    if (mode.startswith("live_") or case.get("network") is True or
            case.get("group") in {"market", "live_options", "live_extraction"}):
        return "LIVE_EXTRACTION"
    return "VBA_CALL"


def is_live(case):
    if case.get("network") is False or case.get("fixture_source") or case.get("domain") == "loopback":
        return False
    return coverage_kind(case) in {"HTTP_ENDPOINT", "LIVE_EXTRACTION"} or (
        coverage_kind(case) == "WORKSHEET" and bool(case.get("domain")))


def causal_http(row, run_id):
    """Require all independently saved case/run/reset-marker fields."""
    if row.get("status", "") not in FAILURES | {"PASS"}:
        return False, "Unfinished/unknown row cannot establish a completed case request"
    if not run_id or row.get("http_observation") != "REQUEST":
        return False, "No case-bound REQUEST observation"
    marker = row.get("http_reset_marker", "")
    prefix = f"smf-test-case:{run_id}:"
    if not marker.startswith(prefix) or not marker[len(prefix):].isdigit():
        return False, "Reset marker does not identify this run/case"
    if row.get("row") and int(marker[len(prefix):]) != int(row["row"]) - 1:
        return False, "Reset marker case number differs from saved row"
    if row.get("http_case_id") != row.get("id") or row.get("http_run_id") != run_id:
        return False, "HTTP case/run ID differs from native row"
    url = row.get("http_url", "")
    try:
        parsed = urlsplit(url)
    except ValueError:
        return False, "Malformed recorded HTTP URL"
    if parsed.scheme.lower() not in {"http", "https"} or not parsed.hostname or url == marker:
        return False, "No real HTTP URL after reset"
    return True, "Request observed after this case's reset marker"


def interpret(row, case, run_id):
    raw = row.get("status", "")
    if raw in UNFINISHED:
        return "UNTESTED", "No completed saved result" + ("; RUNNING was interrupted" if raw == "RUNNING" else "")
    if not case:
        return "UNCLASSIFIED_FAIL", "Saved case has no specification in this suite"
    if raw in {"SETUP_ERROR", "ASSERTION_ERROR", "CHECK_ERROR", "HARNESS_ERROR", "CLEANUP_ERROR"} or (
            raw == "RUNTIME_ERROR" and re.search(r"\[(?:SETUP|CHECK|CLEANUP)\]", row.get("error", ""))):
        return "CODE_FAILURE", "The test setup/assertion/cleanup failed; this does not by itself locate a production defect"
    actual = str(row.get("actual", "")).lower()
    if case.get("expected_obsolete") and (raw == "PASS" or "obsolete" in actual):
        return "OBSOLETE", "The case identifies an obsolete response; it does not establish a working feed"
    if raw == "PASS":
        return "PASS", "This specific assertion passed"
    causal, reason = causal_http(row, run_id)
    try:
        status = int(row.get("http_status", ""))
    except (ValueError, TypeError):
        status = None
    invalid_url = re.search(
        r"(?:\bcurl[^\d\r\n]{0,40}3\b|url[^\r\n]{0,60}malformed|malformed[^\r\n]{0,60}url)",
        row.get("http_error", ""), re.I)
    if is_live(case) and causal and status == 0 and invalid_url:
        return "CODE_FAILURE", "The case supplied or constructed an invalid request URL (same-case curl3/malformed-URL evidence); this is not evidence of a provider outage"
    expected_domain = case.get("domain", "")
    observed_domain = (urlsplit(row.get("http_url", "")).hostname or "").lower() if causal else ""
    domain_matches = (not expected_domain or expected_domain.lower().removeprefix("www.") == observed_domain.removeprefix("www."))
    if is_live(case) and causal and domain_matches:
        if status in UNAVAILABLE_HTTP or (status is not None and 500 <= status <= 599):
            if status == 404:
                return "PROVIDER_UNAVAILABLE", "The requested endpoint returned HTTP404; this does not establish a provider-wide outage or that the requested URL is correct"
            return "PROVIDER_UNAVAILABLE", f"Same-case HTTP request returned {status}"
        if status == 0 and re.search(r"(?:curl[^0-9]*(?:6|7|28|35|56|60)\b|timed? out|could not resolve|failed to connect)", row.get("http_error", ""), re.I):
            return "PROVIDER_UNAVAILABLE", "Same-case HTTP request recorded a transport failure"
        if status == 200 and coverage_kind(case) == "HTTP_ENDPOINT":
            markers = [marker for marker in BLOCKING if marker in actual]
            if markers:
                return "PROVIDER_UNAVAILABLE", "Direct response excerpt contains a rejected access marker: " + ", ".join(markers)
    if raw in FAILURES and not is_live(case):
        return "CODE_FAILURE", "The local/fixture assertion or call failed"
    if causal and not domain_matches:
        reason = "The case-bound request host differs from the case's recorded provider domain"
    return "UNCLASSIFIED_FAIL", "Insufficient evidence to distinguish a code/provider failure; " + reason


def provider_working(row, case):
    expected = str(case.get("expected_text", "")).strip().lower()
    sentinel = expected.startswith(("none", "error", "obsolete")) or bool(case.get("expected_obsolete"))
    return (row["interpreted_status"] == "PASS" and is_live(case)
            and coverage_kind(case) != "HTTP_ENDPOINT" and not sentinel)


def normalized_url(url):
    try:
        parsed = urlsplit(url)
        return parsed._replace(scheme=parsed.scheme.lower(), netloc=parsed.netloc.lower(), fragment="").geturl()
    except ValueError:
        return ""


def select_primary(function, related, specs):
    """Select a recorded direct case without requiring its ID to name the function."""
    for row in related:
        case = specs.get(row["id"], {})
        if case.get("primary") is True or str(case.get("primary_for", "")).lower() == function.lower():
            return row, "Explicit primary case metadata"
    for row in related:
        if row["id"].lower() == function.lower():
            return row, "Case ID matches function name"
    direct = re.compile(r"^\s*(?:[A-Za-z_]\w*\.)?" + re.escape(function) + r"\s*\(", re.I)
    for row in related:
        case = specs.get(row["id"], {})
        if case.get("function", "").lower() == function.lower() and direct.search(case.get("expression", "")):
            return row, "First selected direct-call case; no explicit primary label"
    return None, "No primary or direct-call case identified; indirect coverage is not promoted"


def validate_input_bindings(native, case_path, manifest):
    """Bind saved observations to the exact selected case bytes and source profile."""
    metadata = native.get("run_metadata", {})
    recorded_cases = metadata_value(metadata, "Case specification SHA256", "Case-spec SHA", "Case-spec SHA256")
    manifest_cases = metadata_value(manifest, "case_specification_sha256", "case_spec_sha256")
    actual_cases = hashlib.sha256(Path(case_path).read_bytes()).hexdigest()
    completed = any(row.get("status") in FAILURES | {"PASS"} for row in native["results"])
    if completed and manifest_cases and not recorded_cases:
        raise ValueError("Workbook Case specification SHA256 was not recorded for completed results")
    if recorded_cases and recorded_cases.lower() != actual_cases:
        raise ValueError("Workbook Case specification SHA256 differs from selected suite/test-cases.json (Case-spec SHA mismatch)")
    if manifest_cases and manifest_cases.lower() != actual_cases:
        raise ValueError("Manifest case specification SHA256 differs from selected suite/test-cases.json")
    recorded_source = metadata_value(metadata, "Source fingerprint")
    manifest_source = metadata_value(manifest, "source_fingerprint")
    if completed and manifest_source and not recorded_source:
        raise ValueError("Workbook Source fingerprint was not recorded for completed results")
    if recorded_source and manifest_source and recorded_source.lower() != manifest_source.lower():
        raise ValueError("Workbook Source fingerprint differs from run-manifest.json")
    return {"case_specification_sha256": actual_cases,
            "workbook_case_specification_verified": bool(recorded_cases),
            "manifest_case_specification_verified": bool(manifest_cases),
            "recorded_source_fingerprint": recorded_source,
            "manifest_source_fingerprint": manifest_source,
            "source_fingerprint_verified": bool(recorded_source and manifest_source)}


def build_report(native, specs, inventory, sources, source_meta, manifest=None):
    manifest = manifest or {}
    case_by_id = {case["id"]: case for case in specs}
    if len(case_by_id) != len(specs):
        raise ValueError("Duplicate case IDs in test-cases.json")
    saved = {}
    for row in native["results"]:
        if row["id"] in saved:
            raise ValueError("Duplicate saved result ID: " + row["id"])
        saved[row["id"]] = row
    metadata = native.get("run_metadata", {})
    run_id = metadata_value(metadata, "Run ID")
    expected_version = metadata_value(metadata, "Expected SMF version") or metadata_value(manifest, "expected_version", "expected_smf_version")
    observed_version = metadata_value(metadata, "Actual SMF version", "SMF version")
    issues, warnings = [], list(native.get("diagnostics", []))
    manifest_id = metadata_value(manifest, "run_id")
    if run_id and manifest_id and run_id != manifest_id:
        issues.append("Workbook run ID differs from run-manifest.json")
    if expected_version and observed_version and not version_matches(expected_version, observed_version):
        issues.append("Observed SMF version does not match expected version")
    manifest_version = metadata_value(manifest, "expected_version", "expected_smf_version")
    if manifest_version and expected_version and manifest_version != expected_version:
        issues.append("Workbook expected version differs from run-manifest.json")
    manifest_count = metadata_value(manifest, "case_count", "selected_case_count")
    if manifest_count and manifest_count != str(len(specs)):
        issues.append("Manifest case count differs from test-cases.json")
    if not observed_version:
        warnings.append("Actual SMF version was not recorded; expected version is not proof of the loaded version")
    if not run_id:
        warnings.append("Run ID was not recorded; last-HTTP diagnostics cannot establish current-case availability")
    if not native.get("workbook_vba_sha256"):
        issues.append("Workbook has no saved VBA project")
    rows = []
    for case_id in list(case_by_id) + [case_id for case_id in saved if case_id not in case_by_id]:
        case = case_by_id.get(case_id, {})
        row = dict(saved.get(case_id, {"id": case_id, "status": "MISSING", "function": case.get("function", "")}))
        row["raw_status"] = row.get("status", "")
        row["function"] = row.get("function") or case.get("function", "")
        row["expression"] = case.get("expression", "")
        row["worksheet_formula"] = case.get("worksheet_formula", "")
        row["coverage_kind"] = coverage_kind(case)
        row["interpreted_status"], row["interpretation_reason"] = interpret(row, case, run_id)
        if issues:
            row["interpreted_status"] = "UNTESTED"
            row["interpretation_reason"] = "Run evidence binding is inconsistent: " + "; ".join(issues)
        row["causal_http"], row["http_evidence_note"] = causal_http(row, run_id)
        row["working_provider_case"] = provider_working(row, case)
        row["coverage_note"] = case.get("coverage_note", "")
        rows.append(row)
    public = [item for item in inventory if item.get("visibility", "Public").lower() == "public"
              and item.get("kind", "function").lower() == "function"
              and not str(item.get("file", "")).lower().endswith(".cls")]
    function_rows, gaps = [], []
    source_functions = {ref.get("function", "").lower() for source in sources for ref in source.get("references", [])}
    for item in sorted(public, key=lambda x: x["function"].lower()):
        name = item["function"]
        related = [row for row in rows if row["function"].lower() == name.lower()]
        primary, selection_note = select_primary(name, related, case_by_id)
        entry = {"function": name, "module": item.get("file", ""), "signature": item.get("signature", ""),
                 "primary_case_id": primary["id"] if primary else "",
                 "primary_raw_status": primary["raw_status"] if primary else "MISSING",
                 "primary_interpreted_status": primary["interpreted_status"] if primary else "UNTESTED",
                 "primary_selection": selection_note,
                 "all_case_count": len(related), "raw_counts": dict(Counter(row["raw_status"] for row in related)),
                 "interpreted_counts": dict(Counter(row["interpreted_status"] for row in related)),
                 "case_ids": [row["id"] for row in related],
                 "call_patterns": [{k: row.get(k, "") for k in ("id", "expression", "worksheet_formula", "coverage_kind", "raw_status", "interpreted_status")} for row in related],
                 "working_live_extraction_cases": sum(row["working_provider_case"] for row in related)}
        function_rows.append(entry)
        if not primary:
            gaps.append({"function": name, "pattern": "primary case", "reason": selection_note})
        if not any(row["coverage_kind"] == "WORKSHEET" and row["interpreted_status"] != "UNTESTED" for row in related):
            gaps.append({"function": name, "pattern": "worksheet evaluation", "reason": "No completed worksheet-entry case; VBA calls do not establish worksheet argument/array behavior"})
        if name.lower() in source_functions and not entry["working_live_extraction_cases"]:
            gaps.append({"function": name, "pattern": "live positive extraction", "reason": "No passing positive live extraction for the selected inputs"})
        for row in related:
            if row["interpreted_status"] == "UNTESTED":
                gaps.append({"function": name, "pattern": row["id"], "reason": row["interpretation_reason"]})
    source_rows, endpoint_rows = [], []
    for source in sources:
        domain = source["domain"]
        source_ids = {case["id"] for case in source.get("native_cases", [])}
        related = [row for row in rows if row["id"] in source_ids or case_by_id.get(row["id"], {}).get("domain") == domain]
        endpoints = list(dict.fromkeys(source.get("probe_urls", []) + [source.get("representative_url", "")] +
                                      [item["url"] for item in source.get("extra_probes", []) if item.get("url")]))
        endpoints = [url for url in endpoints if url]
        for endpoint in endpoints:
            observations = [row for row in related if row["causal_http"] and normalized_url(row.get("http_url", "")) == normalized_url(endpoint)]
            endpoint_rows.append({"source": domain, "endpoint": endpoint,
                                  "case_ids": [row["id"] for row in observations],
                                  "request_count": len(observations),
                                  "coverage_status": "OBSERVED" if observations else "UNTESTED",
                                  "raw_statuses": dict(Counter(row["raw_status"] for row in observations)),
                                  "interpreted_statuses": dict(Counter(row["interpreted_status"] for row in observations))})
            if not observations:
                gaps.append({"function": "", "pattern": "source endpoint: " + endpoint,
                             "reason": "No case-bound request to this exact inventoried URL was recorded"})
        source_rows.append({"source_id": source.get("id", domain), "domain": domain, "kind": source.get("kind", ""), "endpoints": endpoints,
                            "case_ids": [row["id"] for row in related],
                            "raw_counts": dict(Counter(row["raw_status"] for row in related)),
                            "interpreted_counts": dict(Counter(row["interpreted_status"] for row in related)),
                            "endpoint_content_passes": sum(row["interpreted_status"] == "PASS" and row["coverage_kind"] == "HTTP_ENDPOINT" for row in related),
                            "positive_extraction_passes": sum(row["working_provider_case"] for row in related),
                            "causal_request_cases": [row["id"] for row in related if row["causal_http"]],
                            "observed_request_hosts": sorted({urlsplit(row["http_url"]).hostname for row in related if row["causal_http"]}),
                            "notes": source.get("notes", "")})
    summary = {"generated_utc": datetime.now(timezone.utc).isoformat(), "run_id": run_id,
               "control_state": metadata_value(metadata, "State"),
               "expected_version": expected_version, "observed_version": observed_version,
               "observed_version_token": observed_version_token(observed_version),
               "version_verified": version_matches(expected_version, observed_version),
               "workbook": native["workbook"], "workbook_sha256": native.get("workbook_sha256"),
               "workbook_vba_sha256": native.get("workbook_vba_sha256"),
               "planned_cases": len(specs), "saved_rows": len(native["results"]),
               "raw_case_counts": dict(Counter(row.get("status", "") for row in native["results"])),
               "planned_case_status_counts": dict(Counter(row["raw_status"] for row in rows)),
               "interpreted_case_counts": dict(Counter(row["interpreted_status"] for row in rows)),
               "public_functions": len(function_rows),
               "primary_raw_counts": dict(Counter(row["primary_raw_status"] for row in function_rows)),
               "primary_interpreted_counts": dict(Counter(row["primary_interpreted_status"] for row in function_rows)),
               "source_groups": len(source_rows), "inventoried_endpoints": len(endpoint_rows),
               "positive_live_extraction_passes": sum(row["working_provider_case"] for row in rows),
               "evidence_errors": issues, "warnings": warnings,
               "scope": "Inventoried public .bas Functions and selected cases only; Subs, Dictionary members, private functions, and every possible argument/provider field are not claimed covered."}
    return {"summary": summary, "functions": function_rows, "sources": source_rows,
            "endpoints": endpoint_rows, "cases": rows, "uncovered_call_patterns": gaps,
            "dynamic_destinations": source_meta.get("dynamic_destinations", []),
            "unreachable_option_references": source_meta.get("unreachable_option_references", [])}


def attach_http_diagnostics(report, directory):
    """Attach separately collected probes; never change a native interpretation."""
    directory = Path(directory).resolve(strict=True)
    raw = load(directory / "results.json")
    metadata = load(directory / "run-metadata.json")
    if not isinstance(raw, list):
        raise ValueError("HTTP results.json must be a list")
    records = []
    for item in raw:
        response = item.get("response", {})
        records.append({"index": item.get("index"), "source_ids": item.get("source_ids", []),
                        "kind": item.get("kind", ""), "mode": item.get("mode", ""),
                        "started_utc": item.get("started_utc", ""), "finished_utc": item.get("finished_utc", ""),
                        "requested_url": response.get("requested_url", ""),
                        "effective_url": response.get("effective_url", ""),
                        "http_status": response.get("http_status", 0),
                        "curl_exit": response.get("curl_exit"), "error": response.get("error", ""),
                        "body_sha256": response.get("body_sha256", ""),
                        "body_path": response.get("body_path", ""), "headers_path": response.get("headers_path", ""),
                        "blocking_markers": response.get("blocking_markers", [])})
    for source in report["sources"]:
        related = [item for item in records if source["source_id"] in item["source_ids"] or source["domain"] in item["source_ids"]]
        source["separate_http_probe_counts"] = dict(Counter(str(item["http_status"]) for item in related))
    report["http_probe_diagnostics"] = {"directory": str(directory), "run_metadata": metadata,
                                         "results_sha256": hashlib.sha256((directory / "results.json").read_bytes()).hexdigest(),
                                         "classification_use": "Diagnostic only; separate probes never override native case/run N:Q evidence or classifications.",
                                         "records": records}
    report["summary"]["separate_http_probe_records"] = len(records)
    report["summary"]["separate_http_status_counts"] = dict(Counter(str(item["http_status"]) for item in records))


def compact(counts):
    return ", ".join(f"{count} {status or 'BLANK'}" for status, count in sorted(counts.items())) or "No cases"


def cell(value, limit=300):
    text = str(value if value is not None else "").replace("|", "\\|").replace("\n", " / ").replace("\r", " ")
    return text if len(text) <= limit else text[:limit - 3] + "..."


def markdown(report):
    s = report["summary"]
    lines = ["# SMF Mac native suite results", "",
             f"Run `{cell(s['run_id'])}` (saved state `{cell(s['control_state'] or 'UNRECORDED')}`) records {s['planned_cases']} planned cases across {s['public_functions']} public functions and {s['source_groups']} inventoried source groups.", "",
             f"Raw results: {compact(s['raw_case_counts'])}. Interpreted results: {compact(s['interpreted_case_counts'])}.", "",
             f"Expected version: `{cell(s['expected_version'])}`. Observed version: `{cell(s['observed_version'] or 'UNRECORDED')}`. Version verified: {s['version_verified']}.", "",
             f"Workbook SHA256: `{s['workbook_sha256']}`. Saved VBA SHA256: `{s['workbook_vba_sha256']}`.", "",
             "A PASS applies to the recorded input/assertion. Many live assertions require only a nonempty or positive result. They do not independently verify the provider's field accuracy, meaning, or freshness. Endpoint-content PASS, offline fixture parsing, error responses, and obsolete responses do not establish a working financial feed.", "",
             "This run uses a fresh checkout on the same Mac. The saved VBA hash identifies the owner workbook, including its imported runner; it does not claim that the packaged add-in binary was tested without modification or that a fresh macOS account was tested.", ""]
    if s.get("input_bindings"):
        binding = s["input_bindings"]
        lines += [f"Case specification SHA256: `{binding['case_specification_sha256']}`. Saved workbook case hash verified: {binding['workbook_case_specification_verified']}. Recorded source fingerprint matches the manifest: {binding['source_fingerprint_verified']}.", ""]
    if s["evidence_errors"] or s["warnings"]:
        lines += ["## Evidence limits", ""] + ["- " + text for text in s["evidence_errors"] + s["warnings"]] + [""]
    lines += ["## Public functions", "",
              "The primary row uses explicit case metadata, then a matching case ID, then the first selected direct call to that function. The CSV records the selection reason and every case; primary status alone does not summarize all behavior.", "",
              "| Function | Primary raw | Primary interpretation | All cases |", "|---|---|---|---|"]
    for row in report["functions"]:
        lines.append(f"| `{cell(row['function'])}` | {row['primary_raw_status']} | {row['primary_interpreted_status']} | {cell(compact(row['interpreted_counts']))} |")
    lines += ["", "## Source groups", "", "| Source | Raw cases | Interpreted cases | Endpoint-content passes | Positive extraction passes |", "|---|---|---|---|---|"]
    for row in report["sources"]:
        lines.append(f"| {cell(row['domain'])} | {cell(compact(row['raw_counts']))} | {cell(compact(row['interpreted_counts']))} | {row['endpoint_content_passes']} | {row['positive_extraction_passes']} |")
    if "http_probe_diagnostics" in report:
        lines += ["", "## Separate HTTP diagnostics", "",
                  f"The explicitly selected HTTP directory contains {s['separate_http_probe_records']} probes: {compact(s['separate_http_status_counts'])} (HTTP status counts). These requests are separate observations; they do not change native case interpretations, and HTTP 200 does not establish content or extraction success.", "",
                  "| Probe | Source IDs | HTTP | curl exit | Requested URL |", "|---|---|---|---|---|"]
        for item in report["http_probe_diagnostics"]["records"]:
            lines.append(f"| {item['index']} | {cell(', '.join(item['source_ids']))} | {item['http_status']} | {item['curl_exit']} | {cell(item['requested_url'], 220)} |")
    lines += ["", "## Cases without a working result", "", "| Case | Raw | Interpretation | Reason / recorded result |", "|---|---|---|---|"]
    for row in report["cases"]:
        if row["interpreted_status"] != "PASS":
            detail = row["interpretation_reason"] + "; " + str(row.get("actual", "")) + "; " + str(row.get("error", ""))
            lines.append(f"| `{cell(row['id'])}` | {row['raw_status'] or 'BLANK'} | {row['interpreted_status']} | {cell(detail, 500)} |")
    lines += ["", "## Coverage limits", "", s["scope"], "",
              f"`uncovered-call-patterns.csv` records {len(report['uncovered_call_patterns'])} missing primary, worksheet, live-extraction, or unfinished call patterns. `report-data.json` lists each function's exact expressions and all source endpoints.", "",
              "Provider-unavailable classifications require matching case/run/reset fields and a same-case HTTP failure or a captured direct-response rejection marker. The label describes the requested endpoint during this run; it does not establish a provider-wide outage or prove that the requested URL is correct. HTTP 200 plus an extraction error remains unclassified when the response content was not captured. Last-HTTP diagnostics alone are not causal evidence.", "",
              "`native-results.private.json` preserves raw native evidence with owner-only file permissions. Shared native/report JSON, Markdown and CSV redact anonymous Yahoo crumbs and cookies; they do not publish those session values.", "",
              "The report reads the saved workbook only. It does not rerun Excel, retry providers, test a fresh Mac, or infer coverage for unbounded caller URLs.", ""]
    return "\n".join(lines)


def write_csv(path, records):
    fields = list(dict.fromkeys(key for row in records for key in row))
    with path.open("w", newline="", encoding="utf-8") as stream:
        writer = csv.DictWriter(stream, fieldnames=fields)
        writer.writeheader()
        for row in records:
            writer.writerow({key: json.dumps(value, ensure_ascii=False) if isinstance(value, (dict, list)) else value for key, value in row.items()})


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--suite", required=True, type=Path)
    parser.add_argument("--workbook", required=True, type=Path)
    parser.add_argument("--output", type=Path, help="Default: SUITE/test-results.md")
    parser.add_argument("--sources", type=Path, help="Default: SUITE/datasource_cases.json")
    parser.add_argument("--source-meta", type=Path, help="Default: SUITE/datasource_inventory_meta.json")
    parser.add_argument("--http-dir", "--httpdir", dest="http_dir", type=Path,
                        help="Optional separate HTTP diagnostics directory; does not classify native failures")
    args = parser.parse_args(argv)
    try:
        suite = args.suite.resolve(strict=True)
        native = read(args.workbook.resolve(strict=True), suite)
        specs = load(suite / "test-cases.json")
        manifest = load(suite / "run-manifest.json") if (suite / "run-manifest.json").exists() else {}
        bindings = validate_input_bindings(native, suite / "test-cases.json", manifest)
        report = build_report(native, specs, load(suite / "function-inventory.json"),
                              load(args.sources or suite / "datasource_cases.json"),
                              load(args.source_meta or suite / "datasource_inventory_meta.json"), manifest)
        report["input_bindings"] = bindings
        report["summary"]["input_bindings"] = bindings
        preparation_path = suite / "preparation.json"
        if preparation_path.exists():
            preparation = load(preparation_path)
            report["preparation_provenance"] = {key: preparation.get(key) for key in
                ("run_id", "expected_version", "source_commit", "source_workbook",
                 "source_workbook_sha256", "source_vba_sha256", "workbook_vba_sha256",
                 "workbook_sha256_before_excel")}
            report["summary"]["source_commit"] = preparation.get("source_commit")
        if args.http_dir:
            attach_http_diagnostics(report, args.http_dir)
        write_private_json(suite / "native-results.private.json", native)
        shared_native = redact_anonymous(native)
        report = redact_anonymous(report)
        output = args.output or suite / "test-results.md"
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(markdown(report), encoding="utf-8")
        (suite / "native-results.json").write_text(json.dumps(shared_native, indent=2) + "\n", encoding="utf-8")
        (suite / "report-data.json").write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
        (suite / "report-summary.json").write_text(json.dumps(report["summary"], indent=2) + "\n", encoding="utf-8")
        for key, filename in [("functions", "function-results.csv"), ("cases", "all-case-results.csv"),
                              ("sources", "datasource-results.csv"), ("endpoints", "endpoint-results.csv"),
                              ("uncovered_call_patterns", "uncovered-call-patterns.csv")]:
            write_csv(suite / filename, report[key])
        if args.http_dir:
            write_csv(suite / "http-probe-results.csv", report["http_probe_diagnostics"]["records"])
    except (OSError, ValueError, KeyError, TypeError, WorkbookReadError) as error:
        parser.exit(2, f"ERROR: {error}\n")
    print(json.dumps(report["summary"], indent=2))
    print(output.resolve())
    return 0


if __name__ == "__main__":
    sys.exit(main())

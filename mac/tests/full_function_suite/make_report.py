#!/usr/bin/env python3
"""Create a report from saved Excel rows and the separate public HTTP probes.

This script reads evidence. It does not run Excel or contact providers.
"""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
import csv
from datetime import datetime, timezone
import json
from pathlib import Path
import re
from urllib.parse import urlsplit

REPO = Path(__file__).resolve().parents[3]
BASE = Path(__file__).resolve().parent
FINAL = {'PASS', 'CHECK_FAILED', 'RUNTIME_ERROR', 'CALL_ERROR', 'CHECK_ERROR', 'SETUP_ERROR', 'HARNESS_ERROR', 'ASSERTION_ERROR'}


def read_json(path):
    return json.loads(Path(path).read_text())


def cell(value, limit=240):
    value = str(value or '').replace('\r', ' ').replace('\n', ' / ').replace('|', '\\|')
    return value if len(value) <= limit else value[:limit-3] + '...'


def phase_status(row):
    status = row.get('status', 'MISSING')
    if status == 'RUNNING':
        return 'INTERRUPTED (saved RUNNING)'
    if status != 'RUNTIME_ERROR':
        return status
    error = row.get('error', '')
    if '[SETUP]' in error:
        return 'SETUP_ERROR'
    if '[CHECK]' in error:
        return 'CHECK_ERROR'
    if '[CALL]' in error:
        return 'CALL_ERROR'
    return 'RUNTIME_ERROR (phase unrecorded)'


def mode(case):
    test_mode = case.get('test_mode', '')
    group = case.get('group', '')
    if test_mode == 'worksheet_comment_image':
        return 'Local picture' if case.get('domain') == 'loopback' else 'Live comment picture'
    if test_mode == 'live_endpoint_request' or group == 'live_datasources':
        return 'Live HTTP/content'
    if test_mode.startswith('live_') or group in ('live_options', 'live_extraction', 'market'):
        return 'Live data extraction'
    if case.get('fixture_source'):
        return 'Offline provider parsing'
    if group == 'options':
        if case.get('function') in ('smfGetOptionExpiry', 'smfGetOptionQuotes'):
            return 'Local option formatting'
        return 'Legacy option path'
    return 'Local/core'


def http_relation(case, row):
    """Only compare explicit case domains; never infer request causality."""
    domain = case.get('domain')
    url = row.get('http_url', '')
    if not domain or not url:
        return 'unverified; may be from an earlier request'
    actual = (urlsplit(url).hostname or '').lower()
    if actual == domain.lower():
        return 'recorded URL matches case domain; may still be cached'
    return 'recorded URL does not match case domain'


def write_csv(path, fields, records):
    with Path(path).open('w', newline='', encoding='utf-8-sig') as stream:
        writer = csv.DictWriter(stream, fieldnames=fields, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(records)


def count_text(rows):
    counts = Counter(r.get('status', 'MISSING') for r in rows)
    return ', '.join(f'{n} {s}' for s, n in sorted(counts.items())) or 'No rows'


def render(native, specs, inventory, sources, assessments, http_records):
    cases = {c['id']: c for c in specs}
    module_by_function = {f['function'].lower(): f.get('file', '') for f in inventory}
    rows = []
    native_rows = list(native['results'])
    present_ids = {r['id'] for r in native_rows}
    for case in specs:
        if case['id'] not in present_ids:
            native_rows.append({'id': case['id'], 'function': case['function'], 'group': case.get('group', ''),
                                'description': case.get('description', ''), 'expected': case.get('expected_text', ''),
                                'status': 'MISSING', 'actual': '', 'error': ''})
    for original in native_rows:
        row = dict(original)
        spec = cases.get(row['id'], {})
        row['mode'] = mode(spec)
        row['reviewed_status'] = phase_status(row)
        row['http_diagnostic_relation'] = http_relation(spec, row)
        row['module'] = spec.get('module', spec.get('source_module', module_by_function.get(row['function'].lower(), '')))
        row['expression'] = spec.get('expression', '')
        row['coverage_note'] = spec.get('coverage_note', '')
        if row['status'] == 'RUNNING':
            row['coverage_note'] += ' No completed result was saved for this attempt; actual/error/HTTP fields may predate the RUNNING checkpoint.'
        rows.append(row)
    by_id = {r['id'].lower(): r for r in rows}
    public = sorted({f['function'] for f in inventory if f['visibility'].lower() == 'public'}, key=str.lower)
    function_rows = []
    for function in public:
        row = dict(by_id.get(function.lower(), {'id': function, 'function': function, 'status': 'MISSING', 'reviewed_status': 'MISSING', 'mode': ''}))
        live = [r for r in rows if r['function'].lower() == function.lower() and r['id'].lower() != function.lower() and r['mode'] in ('Live data extraction', 'Live comment picture')]
        row['other_live_results'] = count_text(live) if live else ''
        function_rows.append(row)
    primary_counts = Counter(r['status'] for r in function_rows)
    completed = [r for r in rows if r['status'] in FINAL]
    pending = [r for r in rows if r['status'] not in FINAL]
    grouped = defaultdict(list)
    for row in rows:
        grouped[row['mode']].append(row)

    source_ids = {}
    for source in sources:
        for c in source.get('native_cases', []):
            source_ids[c['id']] = source['domain']
    native_extract = defaultdict(list)
    for row in rows:
        if row['mode'] not in ('Live data extraction', 'Live comment picture'):
            continue
        spec = cases.get(row['id'], {})
        domain = spec.get('domain') or source_ids.get(row['id'])
        if domain and domain != 'loopback':
            native_extract[domain].append(row)
    source_rows = []
    for a in assessments:
        domain = a['domain']
        related = native_extract[domain]
        counts = Counter(r['status'] for r in related)
        source_rows.append({
            'domain': domain, 'kind': a.get('kind', ''),
            'http_status': a.get('http_status', ''), 'curl_exit': a.get('curl_exit', ''),
            'http_content': 'PASS' if a.get('http_and_content_check_passed') else 'FAILED',
            'url': a.get('url', ''), 'title': a.get('title', ''),
            'blocking_markers': '; '.join(a.get('blocking_markers', [])),
            'native_extraction': count_text(related) if related else 'No extraction case mapped',
            'native_extraction_pass': counts.get('PASS', 0),
            'native_extraction_failed': sum(v for k, v in counts.items() if k in FINAL and k != 'PASS'),
            'native_extraction_pending': sum(v for k, v in counts.items() if k not in FINAL),
            'native_case_ids': '; '.join(r['id'] for r in related),
        })
    source_rows.sort(key=lambda r: r['domain'])
    source_http_pass = sum(a.get('http_and_content_check_passed', False) for a in assessments)
    lines = ['# SMF Mac test results — November 8, 2026', '',
             f'The native evidence records {len(completed)} completed cases. {len(pending)} planned cases are unfinished or missing. {count_text(rows)}.', '',
             f'The suite includes one primary test for each of {len(public)} public functions: {count_text(function_rows)}. Additional cases test indicators, live providers and numbered elements.', '',
             'A PASS applies to the specific input and check in that row. Offline option responses establish parsing against supplied data. Live extraction rows establish whether that call returned the requested data during this run.', '',
             '## Results by test mode', '', '| Test mode | Case results |', '|---|---|']
    for name, values in sorted(grouped.items()):
        lines.append(f'| {cell(name)} | {cell(count_text(values))} |')
    lines += ['', '## Public functions', '',
              'Each row shows the primary test. The last column reports extra live extraction or picture cases for the same function.', '',
              '| Function | Primary result | Test mode | Other live cases |', '|---|---|---|---|']
    for row in function_rows:
        lines.append(f'| `{cell(row["function"], 90)}` | {cell(row["reviewed_status"])} | {cell(row["mode"])} | {cell(row.get("other_live_results", ""))} |')
    lines += ['', '## Data sources', '',
              f'The separate network run made {len(http_records)} requests across {len(assessments)} provider groups. {source_http_pass} provider groups met the HTTP/content checks; {len(assessments)-source_http_pass} did not. The native extraction column reports Excel calls separately.', '',
              'An HTTP 200 response can contain a login page, a parked domain or an unrelated page. HTTP/content checks do not establish that SMF extracted a financial value.', '',
              '| Source | HTTP | HTTP/content check | Native extraction or picture cases |', '|---|---|---|---|']
    for row in source_rows:
        http = str(row['http_status']) if row['http_status'] else f'No response (curl {row["curl_exit"]})'
        lines.append(f'| {cell(row["domain"])} | {cell(http)} | {row["http_content"]} | {cell(row["native_extraction"])} |')
    failed = [r for r in rows if r['status'] in FINAL and r['status'] != 'PASS']
    lines += ['', '## Cases that did not pass', '',
              '`CHECK_FAILED` means the returned value did not meet the expected check. `CALL_ERROR` means the call raised an exception. `CHECK_ERROR` or `ASSERTION_ERROR` means the assertion could not inspect the result. `SETUP_ERROR` means the test could not prepare its inputs. Older runtime rows may have no recorded phase.', '',
              '| Case | Result | Actual value or error |', '|---|---|---|']
    for row in failed:
        actual = row.get('actual', '')
        if row.get('error'): actual += '; ' + row['error']
        lines.append(f'| `{cell(row["id"], 90)}` | {cell(row["reviewed_status"])} | {cell(actual, 360)} |')
    if not failed: lines.append('| — | — | No completed case failed. |')
    if pending:
        lines += ['', '## Unfinished cases', '', 'A saved RUNNING row has no completed result. The interrupted row does not count as a pass or a completed test.', '', '| Case | Saved status |', '|---|---|']
        lines += [f'| `{cell(r["id"], 90)}` | {cell(r["reviewed_status"])} |' for r in pending]
    lines += ['', '## Evidence and limits', '',
              '- `function-results.csv` lists all public functions, the primary results and recorded diagnostics.',
              '- `all-case-results.csv` lists every native row, its call expression, result, error and recorded HTTP fields.',
              '- `datasource-results.csv` lists the separate provider HTTP/content checks and matched native extraction cases.',
              '- `native-results.json` comes from the saved Excel workbook. The HTTP run has separate response bodies, headers, redirect URLs and hashes.',
              '- The last HTTP URL/status can belong to setup, a cached lookup or an earlier case. The report preserves those fields as diagnostics and does not automatically assign a provider failure from them.',
              '- The primary tests and later extra cases ran in batches. Earlier rows can come from an earlier harness revision; saved checkpoints preserve those observations.',
              '- The suite does not directly test every private function, every numbered element or every provider field. A returned price or option quote is a dated observation.', '']
    summary = {'generated_utc': datetime.now(timezone.utc).isoformat(), 'native_case_counts': dict(Counter(r['status'] for r in rows)),
               'primary_public_functions': len(public), 'primary_function_counts': dict(primary_counts),
               'completed_native_cases': len(completed), 'unfinished_native_cases': len(pending),
               'http_requests': len(http_records), 'provider_groups': len(assessments), 'http_content_passed': source_http_pass,
               'mode_counts': {name: dict(Counter(r['status'] for r in values)) for name, values in grouped.items()}}
    return '\n'.join(lines), function_rows, source_rows, rows, summary


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--suite', type=Path)
    parser.add_argument('--http-dir', type=Path, default=REPO/'mac/diagnostics/out/datasource-http-20261108-100559')
    parser.add_argument('--output', type=Path, default=REPO/'mac/TEST-RESULTS-20261108.md')
    args = parser.parse_args()
    suite = args.suite or Path(read_json(REPO/'mac/diagnostics/out/exhaustive-current.json')['suite_dir'])
    native = read_json(suite/'native-results.json')
    specs = read_json(suite/'test-cases.json')
    inventory = read_json(suite/'function-inventory.json')
    sources = read_json(BASE/'datasource_cases.json')
    assessments = read_json(args.http_dir/'provider-assessment.json')
    http_records = read_json(args.http_dir/'results.json')
    report, functions, providers, rows, summary = render(native, specs, inventory, sources, assessments, http_records)
    evidence = [f'- Excel workbook: `{native.get("workbook", "")}`', f'- Native rows: `{suite/"native-results.json"}`', f'- Executed case specifications: `{suite/"test-cases.json"}`', f'- Public HTTP results: `{args.http_dir/"results.json"}`', f'- Provider assessment: `{args.http_dir/"provider-assessment.json"}`']
    notes_path = suite/'run-notes.md'
    if notes_path.exists():
        title, remainder = report.split('\n\n', 1)
        report = title+'\n\n'+notes_path.read_text().strip()+'\n\n'+remainder
    post_fix = suite/'post-fix-notes.md'
    if post_fix.exists():
        title, remainder = report.split('\n\n', 1)
        report = title+'\n\n'+post_fix.read_text().strip()+'\n\n'+remainder
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(report+'\n'+ '\n'.join(evidence)+'\n')
    fields = ['function','id','module','group','mode','status','reviewed_status','other_live_results','expected','actual','error','seconds','http_status','http_url','http_error','http_diagnostic_relation','finished','expression','coverage_note']
    write_csv(suite/'function-results.csv', fields, functions)
    write_csv(suite/'all-case-results.csv', fields, rows)
    write_csv(suite/'datasource-results.csv', ['domain','kind','http_status','curl_exit','http_content','url','title','blocking_markers','native_extraction','native_extraction_pass','native_extraction_failed','native_extraction_pending','native_case_ids'], providers)
    (suite/'report-summary.json').write_text(json.dumps(summary, indent=2)+'\n')
    print(json.dumps(summary, indent=2))
    print(args.output)


if __name__ == '__main__':
    main()

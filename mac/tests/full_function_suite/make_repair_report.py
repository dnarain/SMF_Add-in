#!/usr/bin/env python3
"""Report one saved 0.4 repair run. This script does not run Excel or HTTP calls."""
from __future__ import annotations

import argparse
from collections import Counter, defaultdict
import csv
from datetime import datetime, timezone
import json
from pathlib import Path
import re
from urllib.parse import urlsplit

from make_remaining_report import FINAL, classify, evidence_link, public_text

ROOT = Path(__file__).resolve().parents[3]
EXTERNAL = {
    19902: ('TD Ameritrade', 'Legacy host failed DNS resolution.'),
    702: ('Barchart opinion', 'HTTP 202, empty body, WAF challenge.'),
    995: ('ETFInvestmentOutlook', 'Redirect to an unrelated accounting-services website.'),
    1084: ('ETFScreen', 'Domain-parking page instead of the ETF table.'),
    13626: ('Reuters', 'Legacy and current routes returned HTTP 401 with DataDome protection.'),
    20000: ('MSN StockScouter', 'Legacy host failed DNS resolution.'),
    1116: ('Barchart momentum API', 'Connection to port 443 timed out.'),
}
GROUP_ORDER = ['Regressions and offline fixtures', 'Live data checks',
               'Previously blocked provider checks', 'Image and chart checks', 'Other cases']
SCALAR_ERROR = re.compile(r'^\s*(?:Excel error\b|(?:String:\s*)?(?:Error\b|Not found\b|N/A\b|None\b))', re.I)
HTML_RESPONSE = re.compile(r'<\s*(?:!doctype|/?(?:html|head|script|style|body|table|thead|tbody|tr|td|th|div|span|meta|link))\b', re.I)
USER_PATH = re.compile(r'/Users/[^/\s"<>]+|Macintosh HD:Users:[^:\s"<>]+|[A-Za-z]:\\Users\\[^\\\s"<>]+')


def read_json(path):
    return json.loads(Path(path).read_text())


def safe_text(value, limit=400):
    text = USER_PATH.sub('<user-home>', public_text(value if value is not None else ''))
    text = text.replace('\r', ' ').replace('\n', ' / ').replace('|', '\\|')
    return text if len(text) <= limit else text[:limit - 3] + '...'


def category(case):
    if case.get('element_id') in EXTERNAL:
        return 'Previously blocked provider checks'
    if case.get('group') == 'worksheet_mac_images' or case.get('test_mode', '').startswith('worksheet_mac_'):
        return 'Image and chart checks'
    if case.get('group') in ('json_and_utils', 'json_core', 'repair_local', 'repair_provider_fixture'):
        return 'Regressions and offline fixtures'
    if case.get('group') in ('repair_live_history', 'live_extraction', 'live_options'):
        return 'Live data checks'
    return 'Other cases'


def requested_source(case):
    definition = case.get('source_reference', {}).get('definition', '')
    match = re.search(r'(https?://[^";\s]+)', definition)
    return match[1].replace('~~~~~', case.get('ticker', '')) if match else ''


def make_rows(native, specs):
    if len({c['id'] for c in specs}) != len(specs):
        raise ValueError('Duplicate selected case IDs')
    observed = {}
    for row in native.get('results', []):
        if row['id'] in observed:
            raise ValueError('Duplicate native result ID: ' + row['id'])
        observed[row['id']] = row
    rows = []
    for case in specs:
        row = dict(observed.get(case['id'], {'id': case['id'], 'status': 'UNRUN'}))
        row.setdefault('status', 'UNRUN')
        row['category'] = category(case)
        row['function'] = case['function']
        row['group'] = case.get('group', '')
        row['test_mode'] = case.get('test_mode', '')
        row['description'] = case.get('description', '')
        row['expected'] = row.get('expected') or case.get('expected_text', '')
        row['expression'] = case['expression']
        row['element_id'] = case.get('element_id', '')
        row['source_url'] = requested_source(case)
        row['domain'] = urlsplit(row['source_url']).hostname or case.get('domain', '')
        row['reviewed_status'] = classify(row)
        row['review_note'] = ''
        error = str(row.get('error', ''))
        phase = re.search(r'\[(SETUP|CHECK|CALL)\]', error)
        row['error_phase'] = phase[1] if phase else {
            'SETUP_ERROR': 'SETUP', 'ASSERTION_ERROR': 'CHECK', 'CHECK_ERROR': 'CHECK', 'CALL_ERROR': 'CALL',
        }.get(row['reviewed_status'], '')
        row['recorded_details'] = str(row.get('result_details', row.get('details', '')))
        detail = str(row.get('actual', '')) + '; ' + row['recorded_details']
        row['shape_type'] = (re.search(r'Mac Shape\.Type\s*=\s*(-?\d+)', detail, re.I) or ['', ''])[1]
        row['shape_visible'] = (re.search(r'\bVisible\s*=\s*(-?\d+)', detail, re.I) or ['', ''])[1]
        row['series_values'] = (re.search(r'\bSeries\s*=\s*([^;\r\n]+)', detail, re.I) or ['', ''])[1]
        required_shape = re.search(r'\bpictureType\s*=\s*mso(Picture|Chart)', case.get('check', ''), re.I)
        expected_shape = {'picture': '13', 'chart': '3'}.get(required_shape[1].lower()) if required_shape else None
        if row['status'] == 'PASS':
            if error.strip():
                row['reviewed_status'] = 'INCONSISTENT_PASS'
                row['review_note'] = 'Saved PASS conflicts with the recorded exception.'
            elif row['category'] in ('Live data checks', 'Previously blocked provider checks') and SCALAR_ERROR.match(str(row.get('actual', ''))):
                row['reviewed_status'] = 'INCONSISTENT_PASS'
                row['review_note'] = 'A live data case saved PASS with a scalar error or missing-data value.'
            elif row['category'] == 'Image and chart checks' and expected_shape:
                if row['shape_type'] == '':
                    row['reviewed_status'] = 'UNVERIFIED_IMAGE_PASS'
                    row['review_note'] = 'The harness saved PASS without a recorded object type.'
                elif row['shape_type'] != expected_shape:
                    row['reviewed_status'] = 'INCONSISTENT_PASS'
                    row['review_note'] = 'Saved PASS conflicts with the recorded object type.'
        if row['status'] not in FINAL:
            row['review_note'] = 'No completed result was saved for this attempt.'
        if row['element_id'] in EXTERNAL:
            row['prior_external_observation'] = EXTERNAL[row['element_id']][1]
            row['source_repair_claimed'] = False
        else:
            row['prior_external_observation'] = ''
            row['source_repair_claimed'] = row['element_id'] in (2727, 849, 19004, 5196)
        rows.append(row)
    return rows, sorted(set(observed) - {c['id'] for c in specs})


def summary_for(rows):
    return {
        'planned': len(rows), 'final_statuses': sum(r['status'] in FINAL for r in rows),
        'passes': sum(r['reviewed_status'] == 'PASS' for r in rows),
        'unfinished': sum(r['status'] not in FINAL for r in rows),
        'recorded_status_counts': dict(Counter(r['status'] for r in rows)),
        'reviewed_status_counts': dict(Counter(r['reviewed_status'] for r in rows)),
    }


def count_text(counts):
    return ', '.join(f'{number} {status}' for status, number in sorted(counts.items())) or 'No cases'


def displayed_actual(row):
    if row['status'] not in FINAL:
        return 'No completed result recorded.'
    actual = str(row.get('actual', ''))
    if HTML_RESPONSE.search(actual):
        actual = 'HTML response retained in the local native evidence.'
    pieces = [actual]
    details = row.get('recorded_details', '')
    if details and details not in actual:
        pieces.append('HTML details retained in the local native evidence.' if HTML_RESPONSE.search(details) else details)
    if row.get('error'):
        pieces.append(row['error'])
    if row['review_note']:
        pieces.append(row['review_note'])
    return '; '.join(part for part in pieces if part) or 'No result summary recorded.'


def render(native, specs, suite):
    rows, unexpected = make_rows(native, specs)
    grouped = defaultdict(list)
    for row in rows:
        grouped[row['category']].append(row)
    summaries = {group: summary_for(grouped[group]) for group in GROUP_ORDER if grouped[group]}
    overall = summary_for(rows)
    summary = {
        'generated_utc': datetime.now(timezone.utc).isoformat(), 'suite_dir': str(suite),
        'workbook': native.get('workbook', ''), 'native_control': native.get('control', {}),
        'expected_selected_case_count': 59, 'all_selected_cases_have_final_statuses': overall['unfinished'] == 0,
        'overall': overall, 'categories': summaries, 'unexpected_native_ids': unexpected,
        'seven_external_fields': sorted(EXTERNAL),
        'changed_recipe_live_results': summary_for([r for r in rows if r['source_repair_claimed']]),
    }
    lines = ['# Data and picture repair results — November 8, 2026', '',
             'This report records the saved test-workbook results for the `0.4-dev` source changes.',
             'Add-in build and loading results require separate evidence.', '',
             f'The selected run contains {len(rows)} cases. The harness recorded final statuses for {overall["final_statuses"]}.',
             f'{overall["passes"]} cases passed. {overall["unfinished"]} cases remain unfinished.',
             'Recorded results: ' + count_text(overall['reviewed_status_counts']) + '.', '',
             'The case count includes grouped self-tests. Their internal assertion counts are not added to the case count.',
             'Some negative regressions expect `Error`, `None` or rejection. Their expected behavior appears in the result tables.', '',
             '## Results by group', '', '| Group | Cases | Result statuses |', '| --- | ---: | --- |']
    for group, record in summaries.items():
        lines.append(f'| {safe_text(group)} | {record["planned"]} | {safe_text(count_text(record["reviewed_status_counts"]))} |')
    lines += ['', '## Seven fields without a source repair', '',
              'The seven definitions below retain their original providers. No source repair is claimed for them.',
              'The independent request observations and this run\'s native outcomes are separate records.', '',
              '| Field | Prior external observation | Native result |', '| --- | --- | --- |']
    for item, (provider, observation) in EXTERNAL.items():
        matching = [r for r in rows if r['element_id'] == item]
        result = count_text(Counter(r['reviewed_status'] for r in matching)) if matching else 'No selected case'
        lines.append(f'| {item} — {safe_text(provider)} | {safe_text(observation)} | {safe_text(result)} |')
    lines += ['', 'An observed PASS can establish a successful request for that case. It does not establish a code repair for a prior external blocker.', '']
    for group in GROUP_ORDER:
        if not grouped[group]:
            continue
        lines += ['## ' + group, '', '| Case | Status | Expected result | Returned value and diagnostics |',
                  '| --- | --- | --- | --- |']
        for row in grouped[group]:
            lines.append('| `' + safe_text(row['id'], 100) + '` | ' + safe_text(row['reviewed_status'], 80)
                         + ' | ' + safe_text(row['expected'], 230) + ' | ' + safe_text(displayed_actual(row), 500) + ' |')
        lines.append('')
    lines += ['## Scope and evidence', '',
              '- A PASS applies to the selected inputs and expected result. It does not verify every field on a provider.',
              '- The image cases check returned text, object type, visibility, metadata or plotted values according to each case.',
              '- A picture object does not establish the chart ticker, date, prices or quote freshness.',
              '- The last HTTP URL/status can belong to an earlier request. The local CSV preserves those fields as diagnostics.',
              '- `SETUP_ERROR` records preparation failure. `ASSERTION_ERROR` records a failure inside the result check.',
              '- `CALL_ERROR` records a call-stage exception. The recorded error identifies the failing operation.',
              '- Saved RUNNING, PENDING and missing rows have no completed result. They do not count as passes.',
              '- Windows runtime verification remains separate and pending.',
              '- Historical native rows and earlier reports remain unchanged.', '',
              'The full evidence files remain local under `mac/diagnostics/out` and are excluded from Git.',
              'The public report redacts cookie/crumb values and user paths. It omits raw HTML responses.', '',
              '- ' + evidence_link('Native result rows', suite / 'native-results.json'),
              '- ' + evidence_link('Selected case specifications', suite / 'test-cases.json'),
              '- ' + evidence_link('Full result CSV', suite / 'repair-results.csv'),
              '- ' + evidence_link('Run summary', suite / 'repair-summary.json'),
              '- [Source repair notes](DATA-PICTURE-REPAIRS-20261108.md)', '']
    if native.get('workbook'):
        lines += ['- ' + evidence_link('Saved test workbook', native['workbook']), '']
    if len(rows) != 59:
        lines += [f'The specification contains {len(rows)} cases instead of the initially prepared 59. Counts follow the saved specification.', '']
    if unexpected:
        lines += ['Saved IDs outside the selected specification are excluded: ' + ', '.join(safe_text(x) for x in unexpected) + '.', '']
    report = USER_PATH.sub('<user-home>', public_text('\n'.join(lines), redact_headers=False))
    return report, rows, summary


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--suite', required=True, type=Path, help='Saved run directory containing native-results.json and test-cases.json')
    parser.add_argument('--output', type=Path, default=ROOT / 'mac/DATA-PICTURE-REPAIR-RESULTS-20261108.md')
    args = parser.parse_args()
    native_path = args.suite / 'native-results.json'
    specs_path = args.suite / 'test-cases.json'
    for path in [native_path, specs_path]:
        if not path.exists():
            parser.error('Missing saved evidence: ' + str(path))
    report, rows, summary = render(read_json(native_path), read_json(specs_path), args.suite)
    fields = ['category', 'id', 'function', 'group', 'test_mode', 'element_id', 'domain', 'source_url',
              'status', 'reviewed_status', 'error_phase', 'expected', 'actual', 'shape_type', 'shape_visible',
              'series_values', 'recorded_details', 'error', 'review_note', 'seconds', 'http_status', 'http_url', 'http_error',
              'finished', 'source_repair_claimed', 'prior_external_observation', 'expression']
    with (args.suite / 'repair-results.csv').open('w', newline='', encoding='utf-8-sig') as stream:
        writer = csv.DictWriter(stream, fieldnames=fields, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(rows)
    (args.suite / 'repair-summary.json').write_text(json.dumps(summary, indent=2) + '\n')
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(report)
    print(json.dumps(summary, indent=2))
    print(args.output)


if __name__ == '__main__':
    main()

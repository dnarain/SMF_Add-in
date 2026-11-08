#!/usr/bin/env python3
"""Report the saved remaining-provider run without running Excel or HTTP calls."""
from __future__ import annotations

import argparse
from collections import Counter
import csv
from datetime import datetime, timezone
import json
from pathlib import Path
import re

REPO = Path(__file__).resolve().parents[3]
DEFAULT_SUITE = REPO / 'mac/diagnostics/out/SMF-Mac-remaining-providers-20261108-ew682jmo'
FINAL = {'PASS', 'CHECK_FAILED', 'RUNTIME_ERROR', 'CALL_ERROR', 'CHECK_ERROR',
         'SETUP_ERROR', 'HARNESS_ERROR', 'ASSERTION_ERROR'}
FILL_TYPE = re.compile(r'(?:Comment\s+)?Fill\.Type\s*=\s*(-?\d+)', re.I)
ERROR_VALUE = re.compile(r'^\s*(?:Excel error\b|(?:String:\s*)?(?:Error\b|Not found\b|N/A\b|None\b))', re.I)
SENSITIVE_QUERY = re.compile(r'((?:[?&]|&amp;)(?:crumb|cookie|cookies)=)([^&#\s<>"\x27`]+)', re.I)
SENSITIVE_HEADER = re.compile(r'((?:set-cookie|cookie)\s*:\s*)[^\r\n]+', re.I)


def load_json(path):
    return json.loads(Path(path).read_text())


def public_text(value, redact_headers=True):
    """Redact report text without modifying saved native evidence."""
    text = str(value)
    text = SENSITIVE_QUERY.sub(lambda match: match[1] + '[REDACTED]', text)
    if redact_headers:
        text = SENSITIVE_HEADER.sub(lambda match: match[1] + '[REDACTED]', text)
    return text.replace(str(REPO) + '/', '')


def evidence_link(label, path):
    path = Path(path)
    if not path.is_absolute():
        path = REPO / path
    try:
        relative = path.resolve().relative_to(REPO / 'mac').as_posix()
    except ValueError:
        return f'{label}: local file `{path.name}`'
    return f'[{label}]({relative})'


def text_cell(value, limit=350):
    text = public_text(value if value is not None else '').replace('\r', ' ').replace('\n', ' / ').replace('|', '\\|')
    return text if len(text) <= limit else text[:limit - 3] + '...'


def classify(row):
    """Preserve the raw status and expose its recorded error phase."""
    status = row['status']
    if status == 'RUNNING':
        return 'RUNNING — no saved result'
    if status != 'RUNTIME_ERROR':
        return status
    error = row.get('error', '')
    for phase, label in [('SETUP', 'SETUP_ERROR'), ('CHECK', 'ASSERTION_ERROR'), ('CALL', 'CALL_ERROR')]:
        if '[' + phase + ']' in error:
            return label
    return 'RUNTIME_ERROR — phase unrecorded'


def make_rows(native, cases, parent):
    specs = {c['id']: c for c in cases}
    if len(specs) != len(cases):
        raise ValueError('Duplicate IDs in test-cases.json')
    primary = list(parent['pending_ids'])
    supplemental = list(parent.get('supplemental_ids', []))
    if len(set(primary + supplemental)) != len(primary + supplemental):
        raise ValueError('Duplicate or overlapping primary/supplemental IDs')
    selected = primary + supplemental
    if set(selected) - specs.keys():
        raise ValueError('Selected case missing from test-cases.json')
    observed = {}
    for row in native['results']:
        if row['id'] in observed:
            raise ValueError('Duplicate saved result: ' + row['id'])
        observed[row['id']] = row
    output = []
    for case_id in selected:
        spec = specs[case_id]
        row = dict(observed.get(case_id, {'id': case_id, 'status': 'UNRUN'}))
        row.setdefault('status', 'UNRUN')
        row['category'] = 'primary' if case_id in primary else 'supplemental'
        row['function'] = spec['function']
        row['domain'] = spec.get('domain', '')
        row['element_id'] = spec.get('element_id', '')
        row['expression'] = spec['expression']
        row['description'] = spec.get('description', '')
        row['expected'] = row.get('expected') or spec.get('expected_text', '')
        row['coverage_note'] = spec.get('coverage_note', '')
        row['test_mode'] = spec.get('test_mode', '')
        row['recorded_details'] = row.get('result_details', row.get('details', ''))
        actual = row.get('actual', '')
        match = FILL_TYPE.search(actual + '; ' + str(row['recorded_details']))
        row['fill_type'] = int(match[1]) if match else ''
        row['reviewed_status'] = classify(row)
        phase = re.search(r'\[(SETUP|CHECK|CALL)\]', row.get('error', ''))
        row['error_phase'] = phase[1] if phase else {
            'SETUP_ERROR': 'SETUP', 'ASSERTION_ERROR': 'CHECK',
            'CHECK_ERROR': 'CHECK', 'CALL_ERROR': 'CALL',
        }.get(row['reviewed_status'], '')
        row['review_note'] = ''
        if row['status'] == 'PASS':
            if row.get('error', '').strip():
                row['reviewed_status'] = 'INCONSISTENT_PASS'
                row['review_note'] = 'Saved PASS conflicts with the recorded error field. This row does not count as a pass.'
            elif ERROR_VALUE.match(actual):
                row['reviewed_status'] = 'INCONSISTENT_PASS'
                row['review_note'] = 'Saved PASS conflicts with the recorded error value. This row does not count as a pass.'
            elif row['test_mode'] == 'worksheet_comment_image' and row['fill_type'] not in ('', 6):
                row['reviewed_status'] = 'INCONSISTENT_PASS'
                row['review_note'] = 'Saved PASS conflicts with the recorded picture fill type. This row does not count as a pass.'
            elif row['test_mode'] == 'worksheet_comment_image' and row['fill_type'] == '':
                row['reviewed_status'] = 'UNVERIFIED_IMAGE_PASS'
                row['review_note'] = 'The harness saved PASS but recorded no Fill.Type. Picture evidence is incomplete.'
        if row['status'] not in FINAL:
            row['review_note'] = 'No completed result is recorded. Actual/error/HTTP fields cannot establish the outcome of this attempt.'
        output.append(row)
    unexpected = sorted(set(observed) - set(selected))
    return output, unexpected


def group_summary(rows):
    return {
        'planned': len(rows),
        'final_statuses': sum(r['status'] in FINAL for r in rows),
        'passes': sum(r['reviewed_status'] == 'PASS' for r in rows),
        'unfinished': sum(r['status'] not in FINAL for r in rows),
        'inconsistent_passes': sum(r['reviewed_status'] == 'INCONSISTENT_PASS' for r in rows),
        'unverified_image_passes': sum(r['reviewed_status'] == 'UNVERIFIED_IMAGE_PASS' for r in rows),
        'recorded_status_counts': dict(Counter(r['status'] for r in rows)),
        'reviewed_status_counts': dict(Counter(r['reviewed_status'] for r in rows)),
    }


def counts_text(counts):
    return ', '.join(f'{number} {status}' for status, number in sorted(counts.items())) or 'No results'


def actual_text(row):
    if row['status'] not in FINAL:
        return 'No completed result recorded.'
    parts = [str(row.get('actual', ''))]
    if row['recorded_details'] and str(row['recorded_details']) not in parts[0]:
        parts.append(str(row['recorded_details']))
    if row['fill_type'] != '' and not FILL_TYPE.search('; '.join(parts)):
        parts.append('Comment Fill.Type=' + str(row['fill_type']))
    if row.get('error'):
        parts.append(row['error'])
    if row['review_note']:
        parts.append(row['review_note'])
    return '; '.join(part for part in parts if part) or 'No value summary recorded.'


def result_table(rows):
    lines = ['| Case | Provider or control | Result | Returned value and image details |',
             '| --- | --- | --- | --- |']
    for row in rows:
        lines.append('| `' + text_cell(row['id'], 100) + '` | ' + text_cell(row['domain'], 90)
                     + ' | ' + text_cell(row['reviewed_status'], 90) + ' | '
                     + text_cell(actual_text(row), 450) + ' |')
    return lines


def render(native, cases, parent, suite):
    rows, unexpected = make_rows(native, cases, parent)
    primary = [r for r in rows if r['category'] == 'primary']
    supplemental = [r for r in rows if r['category'] == 'supplemental']
    primary_summary = group_summary(primary)
    supplemental_summary = group_summary(supplemental)
    summary = {
        'generated_utc': datetime.now(timezone.utc).isoformat(),
        'suite_dir': str(suite), 'parent_suite': parent.get('parent_suite', ''),
        'core_version': parent.get('core_version', ''),
        'workbook': native.get('workbook', ''), 'native_control': native.get('control', {}),
        'primary': primary_summary, 'supplemental': supplemental_summary,
        'primary_ids': parent['pending_ids'], 'supplemental_ids': parent.get('supplemental_ids', []),
        'all_selected_cases_have_final_statuses': all(r['status'] in FINAL for r in rows),
        'unexpected_saved_result_ids': unexpected,
        'png_crc_valid_recorded_in_metadata': parent.get('png_crc_valid'),
    }
    lines = ['# Remaining provider results — November 8, 2026', '',
             f'This run uses the isolated full-source workbook with core version `{parent.get("core_version", "unrecorded")}`.',
             f'The selected cases include {len(primary)} original pending checks and {len(supplemental)} supplemental local-image control.', '',
             f'The harness recorded final statuses for {primary_summary["final_statuses"]} of {len(primary)} primary cases.',
             f'{primary_summary["passes"]} primary cases passed. {primary_summary["unfinished"]} primary cases remain unfinished.',
             'Primary results: ' + counts_text(primary_summary['reviewed_status_counts']) + '.', '',
             'The report counts a returned `Error` as a failed data result.',
             '`SETUP_ERROR` records an error during preparation. `ASSERTION_ERROR` records an error in the result check.',
             '`CALL_ERROR` records an exception during the call. Unfinished rows do not count as passes.', '',
             '## The original pending checks', '']
    lines.extend(result_table(primary))
    lines += ['', '## Supplemental local-image control', '',
              'This control is separate from the original pending checks.',
              'Supplemental results: ' + counts_text(supplemental_summary['reviewed_status_counts']) + '.', '']
    lines.extend(result_table(supplemental))
    lines += ['', '## Expected values and diagnostics', '',
              '| Case | Expected result | Recorded error |', '| --- | --- | --- |']
    for row in rows:
        lines.append('| `' + text_cell(row['id'], 100) + '` | ' + text_cell(row['expected'], 350)
                     + ' | ' + text_cell(row.get('error', ''), 300) + ' |')
    lines += ['', '## Interpretation limits', '',
              '- The harness requires returned text `Chart` and `Comment.Shape.Fill.Type = 6` for a picture case.',
              '- The harness does not inspect the picture ticker, date, prices or quote freshness.',
              '- Excel loads comment pictures through `Shape.Fill.UserPicture`. Those requests do not update the SMF HTTP diagnostic fields.',
              '- The SMF last HTTP URL, status and error can come from an earlier request. The CSV preserves them as recorded diagnostics.',
              '- A PASS applies to the selected field or function call. This run does not verify every field on that provider.',
              '- The report does not change the original interrupted workbook or its saved results.',
              '- Netflix lists movie `70111470` as *Iron Man 2* (2010). The new expected year follows [Netflix\'s title page](https://www.netflix.com/title/70111470).', '']
    if unexpected:
        lines += ['Saved rows outside the selected IDs are excluded from the counts: ' + ', '.join('`' + x + '`' for x in unexpected) + '.', '']
    if primary_summary['inconsistent_passes'] or supplemental_summary['inconsistent_passes']:
        lines += ['At least one saved PASS conflicts with its recorded value. The report marks that row `INCONSISTENT_PASS`.', '']
    if primary_summary['unverified_image_passes'] or supplemental_summary['unverified_image_passes']:
        lines += ['A saved picture PASS lacks its recorded fill type. The report marks that row `UNVERIFIED_IMAGE_PASS`.', '']
    lines += ['## Saved evidence', '',
              'These evidence files remain local under `mac/diagnostics/out`. They are excluded from Git.',
              'The relative links resolve in a checkout that retains the evidence.', '',
              '- ' + evidence_link('Native result rows', suite / 'native-results.json'),
              '- ' + evidence_link('Executed case specifications', suite / 'test-cases.json'),
              '- ' + evidence_link('Parent run and selected IDs', suite / 'parent-run.json'),
              '- ' + evidence_link('Full result CSV', suite / 'remaining-results.csv'),
              '- ' + evidence_link('Run summary', suite / 'run-summary.json'), '']
    if native.get('workbook'):
        lines.append('- ' + evidence_link('Saved Excel workbook', native['workbook']))
        lines.append('')
    return public_text('\n'.join(lines), redact_headers=False), rows, summary


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--suite', type=Path, default=DEFAULT_SUITE)
    parser.add_argument('--output', type=Path, default=REPO / 'mac/REMAINING-PROVIDER-RESULTS-20261108.md')
    args = parser.parse_args()
    paths = {name: args.suite / name for name in ['native-results.json', 'test-cases.json', 'parent-run.json']}
    missing = [str(path) for path in paths.values() if not path.exists()]
    if missing:
        parser.error('Missing saved evidence: ' + ', '.join(missing))
    report, rows, summary = render(load_json(paths['native-results.json']), load_json(paths['test-cases.json']),
                                   load_json(paths['parent-run.json']), args.suite)
    fields = ['category', 'id', 'function', 'domain', 'element_id', 'status', 'reviewed_status',
              'description', 'expected', 'actual', 'recorded_details', 'fill_type', 'error', 'error_phase', 'review_note',
              'seconds', 'http_status', 'http_url', 'http_error', 'finished', 'expression', 'test_mode', 'coverage_note']
    with (args.suite / 'remaining-results.csv').open('w', newline='', encoding='utf-8-sig') as stream:
        writer = csv.DictWriter(stream, fieldnames=fields, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(rows)
    (args.suite / 'run-summary.json').write_text(json.dumps(summary, indent=2) + '\n')
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(report)
    print(json.dumps(summary, indent=2))
    print(args.output)


if __name__ == '__main__':
    main()

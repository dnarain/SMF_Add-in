#!/usr/bin/env python3
"""Prepare one disposable SMF native suite using checked-in files only.

This command does not open Excel, install the helper, import VBA, start a server,
or fetch provider data. Existing nonempty run directories are never overwritten.
"""
from __future__ import annotations

import argparse
from collections import Counter
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path
import re
import shlex
import shutil
import struct
import tempfile
import uuid
import zipfile
import zlib

EXPECTED_VERSION = '3.0.2024.07.30-mac.0.7-dev'
CASE_FILES = (
    'json_and_utils_cases.json', 'root_cases.json', 'provider_cases.json',
    'options_cases.json', 'repair_cases.json', 'mac_picture_cases.json',
    'derived_history_cases.json', 'history_rowcount_cases.json', 'history_default_cases.json', 'option_quote1_cases.json',
    'rsi_target_cases.json', 'advfn_year_end_cases.json', 'fred_access_cases.json', 'fred_key_redaction_cases.json',
    'not_supported_cases.json', 'not_supported_field_cases.json', 'google_history_url_cases.json',
)
SOURCE_FILES = ('datasource_cases.json', 'datasource_extraction_cases.json',
                'datasource_inventory_meta.json')
HTML = b'<html><body><table><tr><td>Label</td><td>42</td></tr><tr><td>Earnings</td><td>Nov 20 AMC</td></tr></table></body></html>'
CSV = b'Date,Close,Volume\n2026-11-05,100.5,1000\n2026-11-06,102.25,1200\n'
TEXT = b'SMF fixture response'
ELEMENT = b'19998;Fixture;Value;=1+2\n'
ADDIN_TYPE = b'application/vnd.ms-excel.addin.macroEnabled.main+xml'
WORKBOOK_TYPE = b'application/vnd.ms-excel.sheet.macroEnabled.main+xml'
VBA_LITERAL = re.compile(r'"(?:[^"]|"")*"')
ROUTE = re.compile(r'\{\{LOCAL_HTTP_BASE\}\}([^"\s]*)')
ASSET = re.compile(r'\{\{SUITE_DIR\}\}([^"\s]*)')


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def read_json(path: Path):
    return json.loads(path.read_text(encoding='utf-8'))


def write_json(path: Path, value) -> None:
    path.write_text(json.dumps(value, indent=2, ensure_ascii=True) + '\n', encoding='utf-8')


def png_fixture() -> bytes:
    """Return a small valid RGB PNG; it represents no market/provider data."""
    def chunk(kind, data):
        return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data) & 0xffffffff)
    width, height = 32, 16
    rows = b''.join(b'\0' + b''.join(bytes((32, 128, 192) if (x // 8 + y // 8) % 2 else (240, 240, 240))
                                    for x in range(width)) for y in range(height))
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', width, height, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(rows)) + chunk(b'IEND', b''))


def load_cases(base: Path):
    cases = []
    inputs = []
    for name in CASE_FILES:
        path = base / name
        selected = read_json(path)
        if not isinstance(selected, list):
            raise ValueError(name + ' must contain a list')
        cases.extend(selected)
        inputs.append({'file': name, 'sha256': sha256(path.read_bytes()), 'case_count': len(selected)})
    source_path = base / 'datasource_cases.json'
    sources = read_json(source_path)
    source_cases = [case for source in sources for case in source.get('native_cases', [])]
    cases.extend(source_cases)
    inputs.append({'file': source_path.name, 'sha256': sha256(source_path.read_bytes()), 'case_count': len(source_cases)})
    path = base / 'datasource_extraction_cases.json'
    selected = read_json(path)
    cases.extend(selected)
    inputs.append({'file': path.name, 'sha256': sha256(path.read_bytes()), 'case_count': len(selected)})
    ids = [case['id'].lower() for case in cases]
    if len(set(ids)) != len(ids):
        raise ValueError('Reviewed case IDs are duplicated')
    return cases, inputs


def strings(value):
    if isinstance(value, str):
        yield value
    elif isinstance(value, list):
        for part in value:
            yield from strings(part)
    elif isinstance(value, dict):
        for part in value.values():
            yield from strings(part)


def fixture_audit(cases):
    routes, assets, inline = {}, {}, []
    allowed_routes = {'/text', '/html', '/data.csv', '/image.png', '/denied'}
    allowed_assets = {'', '/fixture.png'}
    for case in cases:
        for value in strings(case):
            for route in ROUTE.findall(value):
                route = route.split('?', 1)[0]
                if route not in allowed_routes:
                    raise ValueError('Unprovisioned loopback route in ' + case['id'] + ': ' + route)
                routes.setdefault(route, []).append(case['id'])
            for asset in ASSET.findall(value):
                if asset not in allowed_assets:
                    raise ValueError('Unprovisioned suite asset in ' + case['id'] + ': ' + asset)
                assets.setdefault(asset or '(suite directory)', []).append(case['id'])
        # Cache bodies and JSON snapshots stay in their reviewed VBA setup.
        # Materialize literal bodies for inspection; native cases are unchanged.
        payloads = []
        for statement in case.get('setup', []):
            if '=' not in statement:
                continue
            rhs = statement.split('=', 1)[1].strip()
            if not VBA_LITERAL.fullmatch(rhs):
                continue
            value = rhs[1:-1].replace('""', '"')
            if value.lstrip().startswith(('{', '[', '<')):
                payloads.append(value)
        if payloads:
            inline.append({'case_id': case['id'], 'payloads': payloads,
                           'origin': 'Stored literal from the checked-in case; preparation performed no provider request.'})
    return {'routes': {k: sorted(set(v)) for k, v in sorted(routes.items())},
            'suite_assets': {k: sorted(set(v)) for k, v in sorted(assets.items())},
            'inline_case_payloads': inline}


def derive_workbook(source: Path, target: Path):
    with zipfile.ZipFile(source) as incoming:
        names = incoming.namelist()
        if len(names) != len(set(names)):
            raise ValueError('The packaged add-in contains duplicate ZIP members')
        if incoming.testzip() is not None:
            raise ValueError('The packaged add-in has a damaged ZIP member')
        content = incoming.read('[Content_Types].xml')
        if content.count(ADDIN_TYPE) != 1:
            raise ValueError('Expected exactly one Excel add-in workbook content type')
        before_vba = incoming.read('xl/vbaProject.bin')
        if not before_vba:
            raise ValueError('The packaged add-in has no VBA project')
        with zipfile.ZipFile(target, 'w') as outgoing:
            for entry in incoming.infolist():
                data = incoming.read(entry.filename)
                if entry.filename == '[Content_Types].xml':
                    data = content.replace(ADDIN_TYPE, WORKBOOK_TYPE, 1)
                outgoing.writestr(entry, data)
    with zipfile.ZipFile(source) as incoming, zipfile.ZipFile(target) as outgoing:
        changed = [name for name in incoming.namelist() if incoming.read(name) != outgoing.read(name)]
        if changed != ['[Content_Types].xml']:
            raise ValueError('Disposable workbook changed unexpected ZIP members: ' + repr(changed))
        if outgoing.testzip() is not None or outgoing.read('xl/vbaProject.bin') != before_vba:
            raise ValueError('Disposable workbook failed VBA byte-preservation checks')
    return {'source_workbook_sha256': sha256(source.read_bytes()),
            'source_vba_sha256': sha256(before_vba),
            'workbook_vba_sha256': sha256(before_vba), 'changed_zip_members': changed}


def prepare(root: Path, suite: Path):
    root, suite = root.resolve(), suite.resolve()
    base = root / 'mac/tests/full_function_suite'
    bundle = root / 'mac/built/0.7-dev'
    source = bundle / 'SMF-Mac.xlam'
    required = [root / 'mac/not_supported.json', root / 'mac/NotSupported.csv', bundle / 'NotSupported.csv', source, bundle / 'BUILD-INFO.json', bundle / 'SHA256SUMS',
                bundle / 'SMFMac.applescript', bundle / 'install_helper.command',
                base / 'build_suite.py', base / 'fixture_server.py']
    required += [base / name for name in CASE_FILES + SOURCE_FILES]
    required += [root / 'mac/tests/provider_recipe_fixtures/manifest.json']
    missing = [str(path.relative_to(root)) for path in required if not path.is_file()]
    if missing:
        raise ValueError('Missing checked-in inputs: ' + ', '.join(missing))
    if suite == root or root in suite.parents and any(part in ('built', 'code', 'docs') for part in suite.relative_to(root).parts):
        raise ValueError('Choose a disposable suite directory outside source/package folders')
    if suite.exists() and (not suite.is_dir() or any(suite.iterdir())):
        raise ValueError('Suite directory must be absent or empty; saved runs are never overwritten')
    build = read_json(bundle / 'BUILD-INFO.json')
    if build.get('version') != EXPECTED_VERSION:
        raise ValueError('Checked-in bundle version does not match ' + EXPECTED_VERSION)
    registry_payload = (root / 'mac/not_supported.json').read_bytes()
    status_payload = (root / 'mac/NotSupported.csv').read_bytes()
    if (bundle / 'NotSupported.csv').read_bytes() != status_payload:
        raise ValueError('Bundle/source retirement status table mismatch')
    if build.get('unsupported_registry_sha256') != sha256(registry_payload):
        raise ValueError('Bundle/source retirement registry mismatch')
    checksum_lines = (bundle / 'SHA256SUMS').read_text().splitlines()
    expected_hash = next((line.split()[0] for line in checksum_lines
                          if len(line.split()) == 2 and line.split()[1].lstrip('*') == 'SMF-Mac.xlam'), None)
    if expected_hash != sha256(source.read_bytes()):
        raise ValueError('Checked-in add-in does not match its SHA256SUMS entry')
    definitions = sorted((root / 'docs/Elements').glob('*.txt'))
    if len(definitions) != 14:
        raise ValueError('Expected 14 checked-in definition files')
    for path in definitions:
        if not (bundle / path.name).is_file() or path.read_bytes() != (bundle / path.name).read_bytes():
            raise ValueError('Bundle/source definition mismatch: ' + path.name)
    provider_root = root / 'mac/tests/provider_recipe_fixtures'
    provider_manifest = read_json(provider_root / 'manifest.json')
    provider_files = []
    for case in provider_manifest['cases']:
        name = case['fixture_file']
        if Path(name).name != name or not name.endswith('.html'):
            raise ValueError('Invalid provider fixture filename: ' + name)
        path = provider_root / name
        if not path.is_file() or sha256(path.read_bytes()) != case['fixture_sha256']:
            raise ValueError('Provider fragment does not match its tracked manifest: ' + name)
        provider_files.append(path)
    if len(provider_files) != 8 or len(set(provider_files)) != 8:
        raise ValueError('Expected eight distinct tracked provider fragments')
    cases, case_inputs = load_cases(base)
    audit = fixture_audit(cases)
    if len(cases) != 538:
        raise ValueError('Expected 538 reviewed cases; found ' + str(len(cases)))
    suite.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix='.' + suite.name + '.prepare-', dir=suite.parent) as temporary:
        work = Path(temporary) / 'run'
        work.mkdir()
        workbook_name = 'SMF-Full-Suite.xlsm'
        hashes = derive_workbook(source, work / workbook_name)
        declared_vba = build.get('vba_project_sha256')
        if declared_vba and declared_vba != hashes['source_vba_sha256']:
            raise ValueError('Bundle VBA project does not match BUILD-INFO.json')
        for path in definitions:
            shutil.copy2(path, work / path.name)
        (work / 'not-supported-registry.json').write_bytes(registry_payload)
        (work / 'NotSupported.csv').write_bytes(status_payload)
        (work / 'smf-elements-suite.txt').write_bytes(ELEMENT)
        (work / 'smf-Elements-Location.txt').write_bytes(b'Local\n')
        for name in CASE_FILES + SOURCE_FILES:
            shutil.copy2(base / name, work / name)
        shutil.copy2(bundle / 'SMFMac.applescript', work / 'SMFMac.applescript')
        shutil.copy2(bundle / 'install_helper.command', work / 'install_helper.command')
        fixtures = work / 'fixtures'
        fixtures.mkdir()
        payloads = {'table.html': HTML, 'data.csv': CSV, 'text.txt': TEXT,
                    'denied.txt': b'fixture denied', 'fixture.png': png_fixture()}
        for name, data in payloads.items():
            (fixtures / name).write_bytes(data)
        (work / 'fixture.png').write_bytes(payloads['fixture.png'])
        provider_output = fixtures / 'provider_recipe_fixtures'
        provider_output.mkdir()
        for path in provider_files + [provider_root / 'manifest.json']:
            shutil.copy2(path, provider_output / path.name)
        write_json(fixtures / 'inline-cache-payloads.json', audit['inline_case_payloads'])
        write_json(fixtures / 'fixture-audit.json', audit)
        routes = {'/text': {'file': 'text.txt', 'status': 200, 'content_type': 'text/plain; charset=utf-8'},
                  '/denied': {'file': 'denied.txt', 'status': 403, 'content_type': 'text/plain; charset=utf-8'},
                  '/html': {'file': 'table.html', 'status': 200, 'content_type': 'text/html; charset=utf-8'},
                  '/data.csv': {'file': 'data.csv', 'status': 200, 'content_type': 'text/csv; charset=utf-8'},
                  '/image.png': {'file': 'fixture.png', 'status': 200, 'content_type': 'image/png'}}
        write_json(fixtures / 'routes.json', routes)
        fixture_files = [{'path': str(path.relative_to(work)), 'bytes': path.stat().st_size,
                          'sha256': sha256(path.read_bytes())}
                         for path in sorted(fixtures.rglob('*')) if path.is_file()]
        run_id = 'smf-0.7-' + uuid.uuid4().hex
        commands = {
            'start_server': ['python3', str(base / 'fixture_server.py'), '--root', str(suite / 'fixtures'),
                             '--output', str(suite / 'server-config.json')],
            'build_runner': ['python3', str(base / 'build_suite.py'), '--suite', str(suite),
                             '--server-config', str(suite / 'server-config.json'),
                             '--run-id', run_id, '--expected-version', EXPECTED_VERSION],
            'install_helper_if_needed': ['/bin/zsh', str(suite / 'install_helper.command')],
        }
        record = {'schema_version': 1, 'run_id': run_id,
                  'prepared_utc': datetime.now(timezone.utc).isoformat(),
                  'suite': str(suite), 'repository': str(root), 'expected_version': 'mac.0.7-dev',
                  'expected_full_version': EXPECTED_VERSION, 'source_commit': build.get('source_commit'),
                  'source_workbook': str(source), 'workbook': str(suite / workbook_name), **hashes,
                  'workbook_sha256_before_excel': sha256((work / workbook_name).read_bytes()),
                  'source_build_info': build, 'case_count': len(cases),
                  'support_registry_sha256': sha256(registry_payload), 'status_table_sha256': sha256(status_payload),
                  'case_groups': dict(Counter(case['group'] for case in cases)), 'case_inputs': case_inputs,
                  'definition_files': [{'name': path.name, 'sha256': sha256(path.read_bytes())} for path in definitions],
                  'fixture_files': fixture_files, 'fixture_audit': {k: v for k, v in audit.items() if k != 'inline_case_payloads'},
                  'server_config': str(suite / 'server-config.json'), 'commands': commands,
                  'timezone_requirement': 'America/Los_Angeles for the three reviewed UTC conversion cases',
                  'native_execution_performed': False, 'external_network_performed': False,
                  'live_provider_substitutions': False,
                  'provider_fixture_count': len(list((fixtures / 'provider_recipe_fixtures').glob('*.html')))}
        write_json(work / 'preparation.json', record)
        instructions = ['Disposable native SMF suite; preparation did not run Excel or HTTP.', '',
                        '1. Start the loopback server in a separate terminal:', shlex.join(commands['start_server']), '',
                        '2. After server-config.json exists, generate the reviewed VBA runner:', shlex.join(commands['build_runner']), '',
                        '3. Install the copied helper only if the reviewed helper is not installed:',
                        shlex.join(commands['install_helper_if_needed']), '',
                        '4. Native Excel owner: open ' + workbook_name + ' as a normal workbook (IsAddin=False).',
                        'Import only modSMFFunctionSuite.bas into that workbook; retain its 51 packaged core components.',
                        'Use SMFRunFullSuite to start all cases, or SMFTestStart for the native operator\'s bounded batches.',
                        'Keep the production add-in unloaded during this isolated run.',
                        'The source add-in and all its VBA bytes are unchanged. No provider answer was supplied by preparation.',
                        'Provider/option/history JSON fixtures remain the reviewed inline case setup; their literals are copied under fixtures for inspection.',
                        'The three UTC conversion cases assume America/Los_Angeles. Record the host timezone before native execution.',
                        'Do not use previous ignored run-state files. Read native results from this workbook and this suite explicitly.', '']
        (work / 'PREPARED.txt').write_text('\n'.join(instructions), encoding='utf-8')
        if suite.exists():
            suite.rmdir()  # Validated empty above; never delete a prior run.
        work.rename(suite)
    return record


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--suite', required=True, type=Path, help='New or empty disposable run directory')
    parser.add_argument('--root', type=Path, default=Path(__file__).resolve().parents[3], help='Fresh checkout root')
    args = parser.parse_args()
    try:
        record = prepare(args.root, args.suite)
    except (ValueError, OSError, KeyError, zipfile.BadZipFile) as error:
        parser.error(str(error))
    print(json.dumps({'suite': record['suite'], 'workbook': record['workbook'],
                      'server_config': record['server_config'], 'case_count': record['case_count'],
                      'commands': record['commands']}, indent=2))


if __name__ == '__main__':
    main()

#!/usr/bin/env python3
"""Check portable suite preparation and loopback fixtures without opening Excel."""
from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
import shutil
import struct
import subprocess
import sys
import tempfile
import time
from urllib.error import HTTPError
from urllib.request import urlopen
import zipfile
import zlib


def png_valid(data):
    if not data.startswith(b'\x89PNG\r\n\x1a\n'):
        return False
    offset, kinds, pixels = 8, [], b''
    while offset < len(data):
        size = struct.unpack('>I', data[offset:offset + 4])[0]
        kind = data[offset + 4:offset + 8]
        body = data[offset + 8:offset + 8 + size]
        crc = struct.unpack('>I', data[offset + 8 + size:offset + 12 + size])[0]
        if crc != zlib.crc32(kind + body) & 0xffffffff:
            return False
        if kind == b'IDAT':
            pixels += body
        kinds.append(kind)
        offset += size + 12
    return offset == len(data) and kinds == [b'IHDR', b'IDAT', b'IEND'] and bool(zlib.decompress(pixels))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, default=Path(__file__).resolve().parents[3])
    parser.add_argument('--preparer', type=Path, default=Path(__file__).with_name('prepare_suite.py'))
    parser.add_argument('--server', type=Path, default=Path(__file__).with_name('fixture_server.py'))
    args = parser.parse_args()
    spec = importlib.util.spec_from_file_location('portable_preparer', args.preparer.resolve())
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    checks = 0

    def check(value, message):
        nonlocal checks
        if not value:
            raise AssertionError(message)
        checks += 1

    def rejected(call, fragment):
        try:
            call()
        except ValueError as error:
            check(fragment in str(error), 'Unexpected rejection: ' + str(error))
        else:
            check(False, 'Expected rejection: ' + fragment)

    def start_server(command, output):
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        for _ in range(100):
            if output.exists():
                return process, json.loads(output.read_text())
            if process.poll() is not None:
                raise AssertionError('Server stopped before readiness: ' + process.communicate()[1])
            time.sleep(0.02)
        process.terminate()
        raise AssertionError('Server readiness file was not written')

    with tempfile.TemporaryDirectory(prefix='smf-fresh-checkout-') as temporary:
        temp = Path(temporary)
        root = temp / 'checkout'
        base = root / 'mac/tests/full_function_suite'
        base.mkdir(parents=True)
        bundle = root / 'mac/built/0.7-dev'
        bundle.mkdir(parents=True)
        for name in ['SMF-Mac.xlam', 'BUILD-INFO.json', 'SHA256SUMS', 'SMFMac.applescript', 'install_helper.command', 'NotSupported.csv']:
            shutil.copy2(args.root / 'mac/built/0.7-dev' / name, bundle / name)
        for name in ['not_supported.json', 'NotSupported.csv']:
            shutil.copy2(args.root / 'mac' / name, root / 'mac' / name)
        for definition in (args.root / 'mac/built/0.7-dev').glob('smf-elements-*.txt'):
            shutil.copy2(definition, bundle / definition.name)
        for name in module.CASE_FILES + module.SOURCE_FILES + ('build_suite.py',):
            shutil.copy2(args.root / 'mac/tests/full_function_suite' / name, base / name)
        shutil.copy2(args.preparer, base / 'prepare_suite.py')
        shutil.copy2(args.server, base / 'fixture_server.py')
        shutil.copytree(args.root / 'docs/Elements', root / 'docs/Elements')
        provider_root = root / 'mac/tests/provider_recipe_fixtures'
        provider_root.mkdir(parents=True)
        manifest = json.loads((args.root / 'mac/tests/provider_recipe_fixtures/manifest.json').read_text())
        for name in ['manifest.json'] + [c['fixture_file'] for c in manifest['cases']]:
            shutil.copy2(args.root / 'mac/tests/provider_recipe_fixtures' / name, provider_root / name)
        (provider_root / 'untracked-not-a-fixture.html').write_text('This file must not be copied.')
        check(not (root / '.git').exists() and not (root / 'mac/diagnostics').exists(), 'Fresh-checkout test has ignored state')
        suite = (temp / 'suite with spaces').resolve()
        record = module.prepare(root, suite)
        check(record['case_count'] == 526, 'Wrong reviewed case count')
        check(record['expected_version'] == 'mac.0.7-dev', 'Missing expected version')
        check(record['native_execution_performed'] is False and record['external_network_performed'] is False, 'Preparation claimed execution')
        check(record['live_provider_substitutions'] is False, 'Preparation substituted live data')
        source = bundle / 'SMF-Mac.xlam'
        workbook = Path(record['workbook'])
        with zipfile.ZipFile(source) as first, zipfile.ZipFile(workbook) as second:
            check(first.namelist() == second.namelist(), 'Workbook archive members changed')
            changed = [name for name in first.namelist() if first.read(name) != second.read(name)]
            check(changed == ['[Content_Types].xml'], 'Unexpected workbook content changes')
            check(first.read('xl/vbaProject.bin') == second.read('xl/vbaProject.bin'), 'VBA changed')
            check(module.WORKBOOK_TYPE in second.read('[Content_Types].xml'), 'Workbook content type not set')
            check(module.ADDIN_TYPE in first.read('[Content_Types].xml'), 'Original add-in changed')
        check((suite / 'smf-elements-suite.txt').read_bytes() == module.ELEMENT, 'Element fixture changed')
        check((suite / 'smf-Elements-Location.txt').read_bytes() == b'Local\n', 'Location fixture changed')
        for definition in (root / 'docs/Elements').glob('*.txt'):
            check((suite / definition.name).read_bytes() == definition.read_bytes(), 'Definition mismatch')
        check(png_valid((suite / 'fixture.png').read_bytes()), 'Invalid root PNG fixture')
        check((suite / 'fixture.png').read_bytes() == (suite / 'fixtures/fixture.png').read_bytes(), 'Image route differs from file fixture')
        check((suite / 'fixtures/table.html').read_bytes() == module.HTML, 'HTML fixture differs')
        check((suite / 'fixtures/data.csv').read_bytes() == module.CSV, 'CSV fixture differs')
        check(len(list((suite / 'fixtures/provider_recipe_fixtures').glob('*.html'))) == 8, 'Copied untracked provider fixture')
        check(not (suite / 'fixtures/provider_recipe_fixtures/untracked-not-a-fixture.html').exists(), 'Untracked provider content copied')
        for name in module.CASE_FILES + module.SOURCE_FILES:
            check((suite / name).read_bytes() == (base / name).read_bytes(), 'Case specification changed')
        check((suite / 'fixtures/inline-cache-payloads.json').exists(), 'Inline fixture provenance missing')
        before = (suite / 'preparation.json').read_bytes()
        rejected(lambda: module.prepare(root, suite), 'absent or empty')
        check((suite / 'preparation.json').read_bytes() == before, 'Existing run overwritten')
        rejected(lambda: module.prepare(root, root / 'mac/built/another-run'), 'outside source/package')
        server = args.server.resolve()
        output = suite / 'server-config.json'
        process, state = start_server([sys.executable, str(server), '--root', str(suite / 'fixtures'), '--output', str(output)], output)
        try:
            check(state['host'] == '127.0.0.1' and state['port'] > 0 and state['pid'] == process.pid, 'Invalid loopback state')
            check(state['fixture_root'] == str(suite / 'fixtures'), 'Wrong fixture root')
            for route, expected in [('/text', module.TEXT), ('/html', module.HTML), ('/data.csv', module.CSV), ('/image.png?t=AAPL', (suite / 'fixture.png').read_bytes())]:
                with urlopen(state['base_url'] + route, timeout=2) as response:
                    check(response.status == 200 and response.read() == expected, 'Wrong loopback body: ' + route)
            for route, status in [('/denied', 403), ('/missing', 404), ('/%2e%2e/SMF-Full-Suite.xlsm', 404)]:
                try:
                    urlopen(state['base_url'] + route, timeout=2)
                except HTTPError as error:
                    check(error.code == status, 'Wrong error status: ' + route)
                else:
                    check(False, 'Expected loopback error: ' + route)
        finally:
            process.terminate()
            process.communicate(timeout=5)
        legacy = temp / 'legacy-server.json'
        process, state = start_server([sys.executable, str(server), str(legacy)], legacy)
        try:
            with urlopen(state['base_url'] + '/image.png', timeout=2) as response:
                check(png_valid(response.read()), 'Legacy PNG CRCs invalid')
        finally:
            process.terminate()
            process.communicate(timeout=5)
        invalid = temp / 'invalid-fixtures'
        invalid.mkdir()
        (invalid / 'routes.json').write_text(json.dumps({'/leak': {'file': '../legacy-server.json', 'status': 200, 'content_type': 'text/plain'}}))
        bad = subprocess.run([sys.executable, str(server), '--root', str(invalid), '--output', str(temp / 'must-not-exist.json')], capture_output=True, text=True)
        check(bad.returncode != 0 and not (temp / 'must-not-exist.json').exists(), 'Unsafe fixture file served')
        broken = root / 'mac/tests/full_function_suite/json_and_utils_cases.json'
        saved = broken.read_bytes()
        cases = json.loads(saved)
        cases[0]['setup'].append('unused = "{{SUITE_DIR}}/missing-fixture.json"')
        broken.write_text(json.dumps(cases))
        rejected(lambda: module.prepare(root, temp / 'unprovisioned-run'), 'Unprovisioned suite asset')
        broken.write_bytes(saved)
        # The installed location finds --root from __file__, without cwd or Git.
        run = subprocess.run([sys.executable, str(base / 'prepare_suite.py'), '--suite', str(temp / 'cli-run')], cwd=temp, capture_output=True, text=True)
        check(run.returncode == 0, 'Fresh-checkout default-root CLI failed: ' + run.stderr)
        check(json.loads(run.stdout)['case_count'] == 526, 'CLI case count differs')
    print('PASS: ' + str(checks) + ' portable preparation/loopback checks; no Excel or external network used.')


if __name__ == '__main__':
    main()

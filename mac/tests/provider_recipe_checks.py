#!/usr/bin/env python3
"""Check public provider fragments independently. This does not execute VBA."""
from datetime import datetime
from decimal import Decimal
import hashlib
from html.parser import HTMLParser
import json
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parents[2]
FIXTURES = Path(__file__).with_name('provider_recipe_fixtures')


class Contents(HTMLParser):
    def __init__(self):
        super().__init__()
        self.stack = []
        self.nodes = []

    def handle_starttag(self, tag, attrs):
        self.stack.append([tag, dict(attrs), []])

    def handle_data(self, data):
        for node in self.stack:
            node[2].append(data)

    def handle_endtag(self, tag):
        for index in range(len(self.stack) - 1, -1, -1):
            if self.stack[index][0] == tag:
                node = self.stack[index]
                self.nodes.append((node[0], node[1], ''.join(node[2]).strip()))
                del self.stack[index:]
                break


def independent_value(case, source):
    parser = Contents()
    parser.feed(source)
    item = case['element_id']
    if item == 2727:
        forward = []
        for tag, _, text in parser.nodes:
            match = re.fullmatch(r'([\d.]+)% \(Forward\)', text)
            if tag == 'div' and match:
                forward.append(Decimal(match[1]) / 100)
        assert len(forward) == 1
        return float(forward[0])
    if item == 849:
        rows = [text for tag, _, text in parser.nodes if tag == 'tr']
        assert len(rows) == 2
        cells = [text for tag, _, text in parser.nodes if tag in ('td', 'th')]
        assert cells[0] == 'Average Price Target'
        return float(Decimal(cells[4].replace('$', '').replace(',', '')))
    if item == 19004:
        metadata = [(attrs, text) for tag, attrs, text in parser.nodes
                    if tag == 'div' and attrs.get('data-uia') == 'metadata']
        assert len(metadata) == 1
        first_item = next(text for tag, _, text in parser.nodes if tag == 'li')
        assert re.fullmatch(r'\d{4}', first_item)
        return float(first_item)
    if item == 5196:
        dates = [text[len('Year End '):] for tag, _, text in parser.nodes
                 if tag == 'th' and text.startswith('Year End ')]
        assert len(dates) == 1
        return datetime.strptime(dates[0], '%B %d %Y').date().isoformat()
    raise AssertionError('Unexpected field ' + str(item))


def main():
    manifest = json.loads((FIXTURES / 'manifest.json').read_text())
    cases = manifest['cases']
    assert len(cases) == 8
    definitions = {}
    for file in (ROOT / 'docs/Elements').glob('smf-elements-*.txt'):
        for line in file.read_text().splitlines():
            parts = line.split(';', 1)
            if parts[0].isdigit() and int(parts[0]) in {2727, 849, 19004, 5196}:
                definitions[int(parts[0])] = parts[1]
    specs = json.loads((ROOT / 'mac/tests/full_function_suite/datasource_extraction_cases.json').read_text())
    for spec in specs:
        item = spec.get('element_id')
        if item in definitions:
            assert spec['source_reference']['definition'] == definitions[item]
    native = (ROOT / 'mac/tests/provider_recipe_tests.bas').read_bytes()
    assert native.count(b'\r\n') == native.count(b'\n')
    native = native.decode('ascii')
    assert max(map(len, native.splitlines())) < 1023
    assert 'If checks <> 25 Then' in native
    for index, case in enumerate(cases, 1):
        path = FIXTURES / case['fixture_file']
        data = path.read_bytes()
        assert hashlib.sha256(data).hexdigest() == case['fixture_sha256']
        source = data.decode('ascii')
        assert '<script' not in source.lower()
        assert 'cookie' not in source.lower() and 'crumb' not in source.lower()
        actual = independent_value(case, source)
        if case['expected_type'] == 'date':
            assert actual == case['expected'], case['id']
        else:
            assert abs(actual - case['expected']) < 1e-12, case['id']
        url = re.search(r'"(https?://[^\"]+)"', definitions[case['element_id']])[1]
        url = url.replace('~~~~~', case['ticker'])
        assert url == case['cached_url']
        assert f'urls({index}) = "{url}"' in native
    msft_yield = next(case for case in cases if case['element_id'] == 2727 and case['ticker'] == 'MSFT')
    assert abs(msft_yield['expected'] - 0.0078) < 1e-12
    print('PASS: 8 public fragments, 4 exact recipe references, cache URL keys, hashes and independent field values.')
    print('Native fixture source has CRLF and 25 planned checks. VBA compilation and execution were not performed.')


if __name__ == '__main__':
    main()

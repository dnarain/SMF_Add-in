#!/usr/bin/env python3
"""Read saved native Excel test rows without changing the workbook."""
from collections import Counter
from pathlib import Path
import json
import posixpath
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
      'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

def read(path):
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            shared = [''.join(si.itertext()) for si in ET.fromstring(z.read('xl/sharedStrings.xml'))]
        rels = {r.attrib['Id']: r.attrib['Target'] for r in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
        sheets = {}
        for s in ET.fromstring(z.read('xl/workbook.xml')).find('s:sheets', NS):
            target = rels[s.attrib['{'+NS['r']+'}id']]
            name = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target)
            cells = {}
            for c in ET.fromstring(z.read(name)).findall('.//s:sheetData/s:row/s:c', NS):
                value = c.find('s:v', NS)
                if c.attrib.get('t') == 'inlineStr':
                    text = ''.join(c.find('s:is',NS).itertext())
                else:
                    text = value.text if value is not None else ''
                    if c.attrib.get('t') == 's': text = shared[int(text)]
                cells[c.attrib['r']] = text
            sheets[s.attrib['name']] = cells
        cells = sheets.get('FunctionResults', {})
        records = []
        for row in range(2, 2000):
            if not cells.get('A'+str(row)): continue
            fields = [('id','A'), ('function','B'),('group','C'),('description','D'),('expected','E'),('status','F'),('seconds','G'),('actual','H'),('error','I'),('http_status','J'),('http_url','K'),('http_error','L'),('finished','M')]
            records.append({key: cells.get(col+str(row),'') for key,col in fields})
        return {'workbook':str(path), 'control':sheets.get('TestControl',{}), 'counts':dict(Counter(r['status'] for r in records)), 'results':records}

if __name__ == '__main__':
    repo = Path(__file__).resolve().parents[3]
    suite = Path(json.loads((repo/'mac/diagnostics/out/exhaustive-current.json').read_text())['suite_dir'])
    result = read(suite/'SMF-Mac-Function-Tests.xlsm')
    (suite/'native-results.json').write_text(json.dumps(result,indent=2)+'\n')
    print(json.dumps({'control':result['control'],'counts':result['counts']}))
    for row in result['results']:
        if row['status'] not in ('PENDING', 'PASS'):
            print(json.dumps(row))

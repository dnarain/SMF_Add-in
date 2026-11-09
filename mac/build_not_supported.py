#!/usr/bin/env python3
"""Check or regenerate the Mac retirement registry data and default-field CSV."""
from __future__ import annotations
import argparse,csv,hashlib,io,json
from pathlib import Path
BEGIN="' BEGIN GENERATED UNSUPPORTED SOURCES"
END="' END GENERATED UNSUPPORTED SOURCES"
HEADERS=['ID','Function or field','Mode','Provider','Status','Reason','Replacement','Evidence','Reviewed']

def quoted(value):
    if not isinstance(value,str) or any(ord(c)>127 or c in '\r\n' for c in value):
        raise ValueError('Embedded retirement metadata must be one-line ASCII text')
    return '"'+value.replace('"','""')+'"'

def generated(registry,payload):
    lines=[BEGIN,"' Generated from mac/not_supported.json; run mac/build_not_supported.py --write.",
        "' Registry SHA256: "+hashlib.sha256(payload).hexdigest(),
        'Private Function smfUnsupportedRegistry() As Collection','    Dim rows As Collection',
        '    Set rows = New Collection']
    for entry in registry['entries']:
        fields=[entry['id'],', '.join(entry['functions']),entry['mode'],entry['provider'],registry['status'],
            entry['reason'],entry['replacement'],'; '.join(entry['evidence']),registry['reviewed']]
        lines.append('    rows.Add _\n        Array( _\n            '+', _\n            '.join(map(quoted,fields))+')')
    lines+=['    Set smfUnsupportedRegistry = rows','End Function','','Private Function smfUnsupportedRoutes() As Collection',
        '    Dim rows As Collection','    Set rows = New Collection']
    for rule in registry['url_rules']:
        lines.append('    rows.Add Array('+', '.join(quoted(rule[x])for x in ('id','host','path'))+')')
    lines+=['    Set smfUnsupportedRoutes = rows','End Function',END]
    return '\n'.join(lines)

def definitions(root):
    effective={}
    for suffix in list(range(20))+[20,21,22]:
        path=root/'docs/Elements'/f'smf-elements-{suffix}.txt'
        if not path.is_file():continue
        for line_number,line in enumerate(path.read_text(errors='replace').splitlines(),1):
            parts=line.split(';')
            if len(parts)>3 and parts[0].isdigit():
                effective[int(parts[0])]=(parts,path.name,line_number)
    return effective

def snapshot(root,registry):
    output=io.StringIO();writer=csv.writer(output,lineterminator='\n');writer.writerow(HEADERS)
    for entry in registry['entries']:
        writer.writerow([entry['id'],', '.join(entry['functions']),entry['mode'],entry['provider'],registry['status'],
            entry['reason'],entry['replacement'],'; '.join(entry['evidence']),registry['reviewed']])
    marked={'="'+v+'"'for v in registry['numbered_fields']['unsupported_literals']}
    for number,(parts,name,line)in sorted(definitions(root).items()):
        if parts[3].strip()in marked:
            writer.writerow([f'field-{number}',f'RCHGetElementNumber field {number}: {parts[2]}',
                'Shipped default; current custom definition can replace it',parts[1],registry['status'],
                'The shipped definition explicitly marks this field obsolete.',
                'Supply a replacement definition or use another supported source.',
                f'docs/Elements/{name}:{line}',registry['reviewed']])
    return output.getvalue()

def run(root,write=False):
    registry_path=root/'mac/not_supported.json';payload=registry_path.read_bytes();registry=json.loads(payload)
    if registry['status']!='Not Supported' or registry['excel_error']!='#N/A':raise ValueError('Unexpected retirement display/error values')
    ids=[e['id']for e in registry['entries']]
    if any(not isinstance(value,str)or not value.strip()for value in ids):raise ValueError('Retirement entry IDs must be nonempty strings')
    if len(ids)!=len(set(ids)) or any(r['id']not in ids for r in registry['url_rules']):raise ValueError('Duplicate or unknown retirement ID')
    for rule in registry['url_rules']:
        if not rule['path'].startswith('/') or '?'in rule['path']or '#'in rule['path']or rule['host']!=rule['host'].lower():
            raise ValueError('Retirement URL rules must contain a lowercase host and an exact path')
    code_path=root/'code/modGetElementNumber.bas';code=code_path.read_text()
    if code.count(BEGIN)!=1 or code.count(END)!=1:raise ValueError('Retirement data markers are missing or repeated')
    a,end_at=code.index(BEGIN),code.index(END)
    if a>=end_at:raise ValueError('Retirement data markers are out of order')
    b=end_at+len(END)
    updated=code[:a]+generated(registry,payload)+code[b:]
    csv_path=root/'mac/NotSupported.csv';table=snapshot(root,registry)
    mismatches=[]
    if updated!=code:mismatches.append(str(code_path.relative_to(root)))
    if not csv_path.exists()or csv_path.read_text()!=table:mismatches.append(str(csv_path.relative_to(root)))
    if write:
        if updated!=code:code_path.write_text(updated)
        if not csv_path.exists()or csv_path.read_text()!=table:csv_path.write_text(table)
    elif mismatches:raise ValueError('Generated retirement data differs: '+', '.join(mismatches))
    rows=list(csv.DictReader(io.StringIO(table)))
    return {'registry_sha256':hashlib.sha256(payload).hexdigest(),'source_groups':len(ids),
        'exact_url_rules':len(registry['url_rules']),'default_fields':sum(x['ID'].startswith('field-')for x in rows),
        'status_rows':len(rows),'mode':'write'if write else 'check','result':'PASS'}

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--root',type=Path,default=Path(__file__).resolve().parents[1])
    parser.add_argument('--write',action='store_true',help='Update the embedded registry data and CSV; default checks only')
    args=parser.parse_args()
    try:print(json.dumps(run(args.root.resolve(),args.write),indent=2))
    except (ValueError,KeyError,OSError)as error:parser.exit(1,'ERROR: '+str(error)+'\n')

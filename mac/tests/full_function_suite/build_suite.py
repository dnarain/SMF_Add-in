#!/usr/bin/env python3
"""Build a disposable native Excel test workbook module from reviewed cases.

This script writes source files. It neither imports VBA nor runs Excel.
"""
from pathlib import Path
import argparse
import json
import re
import sys

REPO = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO / 'mac/tests'))
from port_static_checks import active_lines, logical_lines, DECLARATION
BASE = Path(__file__).resolve().parent


def inventory():
    found = []
    for path in sorted((REPO / 'code').glob('*.bas')):
        if path.name == 'exportMods.bas':
            continue
        for number, line in logical_lines(active_lines(path.read_text(), {'mac': 1, 'vba7': -1, 'win64': 0, 'win32': 0})):
            match = DECLARATION.match(line)
            if match and match[2].lower() == 'function':
                found.append({'function': match[3], 'visibility': match[1] or 'Public', 'file': str(path.relative_to(REPO)), 'line': number, 'signature': line})
    return found


def quoted(value):
    return '"' + str(value).replace('"', '""').replace('\r', ' ').replace('\n', ' ') + '"'


HELPERS = '''
Private Function ArrayCell(ByVal value As Variant, ByVal row As Long, ByVal col As Long) As Variant
    On Error GoTo Missing
    If Not IsArray(value) Then Exit Function
    Dim cell As Variant
    If col = 0 Then
        cell = value(row)
    Else
        cell = value(row, col)
    End If
    If IsError(cell) Or IsNull(cell) Or IsEmpty(cell) Then Exit Function
    ArrayCell = cell
Missing:
End Function

Private Function DateNumber(ByVal value As Variant) As Double
    On Error GoTo Missing
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    If IsDate(value) Then DateNumber = CDbl(CDate(value))
Missing:
End Function

Private Function IsUsableMarketResult(ByVal value As Variant) As Boolean
    On Error GoTo Missing
    If IsObject(value) Or IsArray(value) Then Exit Function
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    Dim text As String
    text = LCase$(Trim$(CStr(value)))
    If Len(text) = 0 Then Exit Function
    Select Case text
        Case "error", "n/a", "na", "none", "not found", "premium", "--", "-", "null", "nothing"
            Exit Function
    End Select
    If InStr(text, "error") > 0 Or InStr(text, "not found") > 0 Then Exit Function
    If InStr(text, "<html") > 0 Or InStr(text, "<!doctype") > 0 Then Exit Function
    IsUsableMarketResult = True
Missing:
End Function

Private Function ArrayHasData(ByVal value As Variant) As Boolean
    On Error GoTo Missing
    If Not IsArray(value) Then Exit Function
    Dim part As Variant
    For Each part In value
        If IsUsableMarketResult(part) Then
            If IsNumeric(part) Or IsDate(part) Then
                ArrayHasData = True
                Exit Function
            End If
        End If
    Next part
Missing:
End Function

Private Function FixtureHTML() As String
    FixtureHTML = "<html><body><table><tr><td>Label</td><td>42</td></tr><tr><td>Earnings</td><td>Nov 20 AMC</td></tr></table></body></html>"
End Function

Private Function SummarizeValue(ByVal value As Variant, ByVal obj As Object) As String
    On Error GoTo Missing
    Dim part As Variant, count As Long, text As String
    If Not obj Is Nothing Then
        SummarizeValue = TypeName(obj) & "; count=" & CStr(obj.Count)
    ElseIf IsArray(value) Then
        text = "Array "
        For Each part In value
            count = count + 1
            If count <= 16 Then
                If IsError(part) Then
                    text = text & "[Excel error] | "
                ElseIf IsNull(part) Then
                    text = text & "Null | "
                Else
                    text = text & Left$(CStr(part), 100) & " | "
                End If
            End If
        Next part
        SummarizeValue = text & " cells=" & CStr(count)
    ElseIf IsError(value) Then
        SummarizeValue = "Excel error " & CStr(value)
    ElseIf IsNull(value) Then
        SummarizeValue = "Null"
    ElseIf IsEmpty(value) Then
        SummarizeValue = "Empty"
    Else
        SummarizeValue = TypeName(value) & ": " & Left$(CStr(value), 500)
        If VarType(value) = vbString Then SummarizeValue = SummarizeValue & "; length=" & CStr(Len(value))
    End If
    Exit Function
Missing:
    SummarizeValue = "Result summary unavailable: " & Err.Description
End Function

Private Function SuiteSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set SuiteSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If SuiteSheet Is Nothing Then
        Set SuiteSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        SuiteSheet.Name = name
    End If
End Function

Public Sub SMFTestStop()
    SuiteSheet("TestControl").Range("B3").Value2 = "STOP"
    ThisWorkbook.Save
End Sub

Private Sub ScheduleNext()
    Application.OnTime Now + TimeSerial(0, 0, 1), "'" & ThisWorkbook.Name & "'!SMFTestRunNext"
End Sub
'''


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--suite', type=Path)
    parser.add_argument('--server-config', type=Path, default=REPO / 'mac/diagnostics/out/exhaustive-server.json')
    parser.add_argument('--case-file', type=Path, help='Generate only these reviewed cases; do not require all public functions.')
    args = parser.parse_args()
    suite = args.suite
    if suite is None:
        state = json.loads((REPO / 'mac/diagnostics/out/exhaustive-current.json').read_text())
        suite = Path(state['suite_dir'])
    # Fixture filenames must remain absolute inside Excel's workbook formulas.
    suite = suite.resolve()
    server = json.loads(args.server_config.read_text())
    files = ['json_and_utils_cases.json', 'root_cases.json', 'provider_cases.json', 'options_cases.json',
             'repair_cases.json', 'mac_picture_cases.json', 'derived_history_cases.json']
    missing_files = [name for name in files if not (BASE / name).exists()]
    if missing_files:
        raise SystemExit('Missing reviewed case files: ' + ', '.join(missing_files))
    cases = []
    for name in files:
        cases.extend(json.loads((BASE / name).read_text()))
    if (BASE / 'datasource_cases.json').exists():
        sources = json.loads((BASE / 'datasource_cases.json').read_text())
        for source in sources:
            cases.extend(source.get('native_cases', []))
    if (BASE / 'datasource_extraction_cases.json').exists():
        cases.extend(json.loads((BASE / 'datasource_extraction_cases.json').read_text()))
    if args.case_file:
        cases = json.loads(args.case_file.read_text())
    tokens = {'{{SUITE_DIR}}': str(suite), '{{LOCAL_HTTP_BASE}}': server['base_url'], '{{SUITE_SHEET}}': 'TestInputs'}
    raw = json.dumps(cases)
    for token, replacement in tokens.items():
        raw = raw.replace(token, replacement)
    cases = json.loads(raw)
    assert len({c['id'].lower() for c in cases}) == len(cases), 'Duplicate case IDs'
    functions = inventory()
    public = {f['function'].lower() for f in functions if f['visibility'].lower() == 'public'}
    covered = {c['function'].lower() for c in cases}
    if not args.case_file:
        assert public <= covered, 'Missing function cases: ' + repr(public - covered)
    # Network calls use a small enough bounded request to preserve the saved RUNNING case.
    out = ['Attribute VB_Name = "modSMFFunctionSuite"', 'Option Explicit', f'Private Const SuiteCount As Long = {len(cases)}', HELPERS]
    out += ['Public Sub SMFTestStart(Optional ByVal firstCase As Long = 1, Optional ByVal lastCase As Long = 0)',
            '    Dim ws As Worksheet, inputs As Worksheet, control As Worksheet, i As Long',
            '    Set ws = SuiteSheet("FunctionResults")', '    Set inputs = SuiteSheet("TestInputs")',
            '    Set control = SuiteSheet("TestControl")',
            '    ThisWorkbook.Worksheets(1).Range("A1:B4").ClearContents',
            '    ws.Cells.NumberFormat = "@"',
            '    ws.Range("A1:M1").Value = Array("Case", "Function", "Group", "Description", "Expected", "Status", "Seconds", "Actual", "Error", "Last HTTP status", "Last HTTP URL", "Last HTTP error", "Finished")',
            '    For i = 1 To 5: inputs.Cells(i, 1).Value2 = i: Next i',
            '    inputs.Range("J20").ColumnWidth = 30', '    inputs.Range("J20").RowHeight = 60',
            '    If lastCase = 0 Then lastCase = SuiteCount',
            '    control.Range("A1:A3").Value = Application.Transpose(Array("Next case", "Last case", "State"))',
            '    control.Range("B1").Value2 = firstCase', '    control.Range("B2").Value2 = lastCase',
            '    control.Range("B3").Value2 = "RUN"']
    metadata_start = len(out)
    for i, c in enumerate(cases, 1):
        out += [f'    If Len(ws.Cells({i+1}, 1).Value2) = 0 Then', f'        ws.Cells({i+1}, 1).Value2 = {quoted(c["id"])}',
                f'        ws.Cells({i+1}, 2).Value2 = {quoted(c["function"])}', f'        ws.Cells({i+1}, 3).Value2 = {quoted(c["group"])}',
                f'        ws.Cells({i+1}, 4).Value2 = {quoted(c["description"])}', f'        ws.Cells({i+1}, 5).Value2 = {quoted(c["expected_text"])}',
                f'        ws.Cells({i+1}, 6).Value2 = "PENDING"', '    End If']
    metadata_lines = out[metadata_start:]
    del out[metadata_start:]
    metadata_subs = []
    for offset in range(0, len(cases), 25):
        sub_name = 'InitializeCases' + str(offset // 25 + 1)
        out.append('    ' + sub_name + ' ws')
        metadata_subs += ['Private Sub ' + sub_name + '(ByVal ws As Worksheet)', *metadata_lines[offset*8:(offset+25)*8], 'End Sub']
    out += ['    ws.Columns("A:C").AutoFit', '    ThisWorkbook.Save', '    ScheduleNext', 'End Sub',
            'Public Sub SMFTestRunNext()',
            '    Dim ws As Worksheet, control As Worksheet, n As Long, row As Long, started As Double',
            '    Dim result As Variant, obj As Object, ok As Boolean, detail As String',
            '    Set control = SuiteSheet("TestControl")', '    If control.Range("B3").Value2 <> "RUN" Then Exit Sub',
            '    n = CLng(control.Range("B1").Value2)',
            '    If n > CLng(control.Range("B2").Value2) Or n > SuiteCount Then',
            '        control.Range("B3").Value2 = "DONE"', '        ThisWorkbook.Save', '        Exit Sub', '    End If',
            '    Set ws = SuiteSheet("FunctionResults")', '    row = n + 1',
            '    ws.Range(ws.Cells(row, 7), ws.Cells(row, 13)).ClearContents',
            '    ws.Cells(row, 6).Value2 = "RUNNING"', '    ThisWorkbook.Save',
            '    started = Timer', '    Application.StatusBar = "SMF test " & CStr(n) & "/" & CStr(SuiteCount) & ": " & ws.Cells(row, 1).Value2',
            '    On Error GoTo Failed', '    Select Case n']
    for i, c in enumerate(cases, 1):
        out.append(f'        Case {i}: SuiteCase{i:03d} result, obj, ok, detail')
    out += ['    End Select', '    If ok Then', '        ws.Cells(row, 6).Value2 = "PASS"', '    Else',
            '        ws.Cells(row, 6).Value2 = "CHECK_FAILED"', '    End If', '    GoTo Finished',
            'Failed:', '    ws.Cells(row, 6).Value2 = "RUNTIME_ERROR"',
            '    If InStr(Err.Source, "[CHECK]") > 0 Then ws.Cells(row, 6).Value2 = "ASSERTION_ERROR"',
            '    If InStr(Err.Source, "[SETUP]") > 0 Then ws.Cells(row, 6).Value2 = "SETUP_ERROR"',
            '    ws.Cells(row, 9).Value2 = CStr(Err.Number) & ": " & Err.Source & ": " & Err.Description', '    Resume Finished',
            'Finished:', '    On Error Resume Next',
            '    ws.Cells(row, 7).Value2 = CStr((Timer - started + 86400#) - 86400# * Int((Timer - started + 86400#) / 86400#))',
            '    ws.Cells(row, 8).Value2 = Left$(SummarizeValue(result, obj) & IIf(Len(detail) > 0, "; " & detail, ""), 2000)',
            '    ws.Cells(row, 10).Value2 = CStr(SMFMacHTTPStatus())',
            '    ws.Cells(row, 11).Value2 = SMFMacLastURL()', '    ws.Cells(row, 12).Value2 = SMFMacLastError()',
            '    ws.Cells(row, 13).Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")',
            '    control.Range("B1").Value2 = n + 1', '    ThisWorkbook.Save', '    On Error GoTo 0',
            '    If n < CLng(control.Range("B2").Value2) Then', '        ScheduleNext', '    Else',
            '        control.Range("B3").Value2 = "DONE"', '        ThisWorkbook.Save',
            '        Application.StatusBar = "SMF test batch finished"', '    End If', 'End Sub']
    out.extend(metadata_subs)
    for i, c in enumerate(cases, 1):
        out += [f'Private Sub SuiteCase{i:03d}(ByRef result As Variant, ByRef obj As Object, ByRef ok As Boolean, ByRef detail As String)',
                '    Dim savedNumber As Long, savedText As String, phase As String', *['    '+x for x in c.get('declarations', [])],
                '    phase = "SETUP"', '    On Error GoTo Failed', *['    '+x for x in c.get('setup', [])]]
        out.append('    phase = "CALL"')
        out.extend('    ' + step for step in c.get('call_steps', []))
        out.append('    '+ ('Set obj = ' if c.get('object_result') else 'result = ') + c['expression'])
        out.append('    phase = "CHECK"')
        out.extend('    ' + step for step in c.get('check_setup', []))
        if c.get('result_details'):
            out.append('    detail = ' + c['result_details'])
        out.append('    ok = ' + c['check'])
        out += ['    GoTo CleanUp', 'Failed:', '    savedNumber = Err.Number', '    savedText = Err.Description', '    Resume CleanUp',
                'CleanUp:', '    On Error Resume Next', *['    '+x for x in c.get('cleanup', [])], '    On Error GoTo 0',
                f'    If savedNumber <> 0 Then Err.Raise savedNumber, {quoted(c["id"])} & " [" & phase & "]", savedText', 'End Sub']
    source = '\r\n'.join(out) + '\r\n'
    assert max(map(len, source.splitlines())) <= 1023, 'VBA physical line exceeds1023characters'
    assert '{{' not in source, 'Unresolved fixture placeholder'
    (suite/'modSMFFunctionSuite.bas').write_bytes(source.encode('utf-8'))
    (suite/'test-cases.json').write_text(json.dumps(cases,indent=2)+'\n')
    (suite/'function-inventory.json').write_text(json.dumps(functions,indent=2)+'\n')
    print(f'{len(public)} public functions in source; {len(cases)} selected cases; full function coverage required: {not bool(args.case_file)}')
    print(suite/'modSMFFunctionSuite.bas')

if __name__ == '__main__':
    main()

#!/usr/bin/env python3
"""Generate reviewed native Excel tests without importing VBA or running Excel.

Every run uses explicit output and loopback-fixture arguments. Generated metadata
describes the selected source/specifications; it does not establish runtime results.
"""
from pathlib import Path
import argparse
from datetime import datetime, timezone
import hashlib
import json
import re
import sys
from urllib.parse import urlsplit
from uuid import uuid4

REPO = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO / "mac/tests"))
from port_static_checks import active_lines, logical_lines, DECLARATION
BASE = Path(__file__).resolve().parent
EXPECTED_VERSION = "3.0.2024.07.30-mac.0.6-dev"
DEFAULT_CASE_FILES = (
    "json_and_utils_cases.json", "root_cases.json", "provider_cases.json", "options_cases.json",
    "repair_cases.json", "mac_picture_cases.json", "derived_history_cases.json",
    "history_rowcount_cases.json", "history_default_cases.json", "option_quote1_cases.json",
)


def inventory():
    found = []
    for path in sorted((REPO / "code").glob("*.bas")):
        if path.name == "exportMods.bas":
            continue
        constants = {"mac": 1, "vba7": -1, "win64": 0, "win32": 0}
        for number, line in logical_lines(active_lines(path.read_text(), constants)):
            match = DECLARATION.match(line)
            if match and match[2].lower() == "function":
                found.append({"function": match[3], "visibility": match[1] or "Public",
                              "file": str(path.relative_to(REPO)), "line": number, "signature": line})
    return found


def quoted(value):
    return '"' + str(value).replace('"', '""').replace("\r", " ").replace("\n", " ") + '"'


def digest(payload):
    return hashlib.sha256(payload).hexdigest()


def local_base(value):
    value = value.strip().rstrip("/")
    parsed = urlsplit(value)
    if (parsed.scheme not in {"http", "https"}
            or parsed.hostname not in {"127.0.0.1", "localhost", "::1"}
            or parsed.username is not None or parsed.password is not None
            or parsed.query or parsed.fragment or any(c in value for c in '\r\n"')):
        raise ValueError("The fixture base must be an explicit HTTP(S) loopback URL")
    _ = parsed.port  # Reject malformed or out-of-range ports without connecting.
    return value


def replace_tokens(value, tokens):
    if isinstance(value, str):
        for token, replacement in tokens.items():
            value = value.replace(token, replacement)
        return value
    if isinstance(value, list):
        return [replace_tokens(item, tokens) for item in value]
    if isinstance(value, dict):
        return {key: replace_tokens(item, tokens) for key, item in value.items()}
    return value


def load_cases(case_file):
    inputs = []
    def read(path):
        payload = path.read_bytes()
        inputs.append({"file": str(path.resolve()), "sha256": digest(payload)})
        return json.loads(payload)
    if case_file:
        return read(case_file), inputs
    cases = []
    for name in DEFAULT_CASE_FILES:
        cases.extend(read(BASE / name))
    for provider in read(BASE / "datasource_cases.json"):
        cases.extend(provider.get("native_cases", []))
    cases.extend(read(BASE / "datasource_extraction_cases.json"))
    return cases, inputs


def validate_cases(cases, public, require_all):
    if not isinstance(cases, list) or not cases:
        raise ValueError("The case file must contain a nonempty JSON array")
    required = {"id", "function", "group", "description", "expression", "check", "expected_text"}
    for case in cases:
        missing = required - case.keys()
        if missing:
            raise ValueError(f"Case is missing fields: {sorted(missing)}")
        for field in required:
            if not isinstance(case[field], str) or not case[field]:
                raise ValueError(f"Case field {field!r} must be nonempty text")
        for field in ("declarations", "setup", "call_steps", "check_setup", "cleanup"):
            if not isinstance(case.get(field, []), list) or not all(isinstance(x, str) for x in case.get(field, [])):
                raise ValueError(f"{case['id']}: {field} must contain VBA source strings")
    if len({case["id"].casefold() for case in cases}) != len(cases):
        raise ValueError("Duplicate case IDs")
    covered = {case["function"].lower() for case in cases}
    if require_all and not public <= covered:
        raise ValueError("Missing public function cases: " + repr(sorted(public - covered)))


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


Private Sub CaptureStatus()
    If Not mStatusCaptured Then
        mSavedStatus = Application.StatusBar
        mStatusCaptured = True
    End If
End Sub

Private Sub RestoreStatus()
    If mStatusCaptured Then Application.StatusBar = mSavedStatus
    mStatusCaptured = False
End Sub

Private Function RunMacroName() As String
    RunMacroName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!SMFTestRunNext"
End Function

Private Sub CancelScheduled()
    On Error Resume Next
    If mScheduled Then Application.OnTime mNextRun, RunMacroName(), Schedule:=False
    mScheduled = False
End Sub

Private Sub ScheduleNext()
    CancelScheduled
    mNextRun = Now + TimeSerial(0, 0, 1)
    Application.OnTime mNextRun, RunMacroName()
    mScheduled = True
End Sub

Private Sub SuiteJournal(ByVal pCase As Long, ByVal pState As String, ByVal pDetail As String)
    Dim ws As Worksheet, rowNumber As Long
    Set ws = SuiteSheet("TestJournal")
    If Len(ws.Range("A1").Value2) = 0 Then
        ws.Range("A1:F1").Value = Array("Timestamp", "Case", "ID", "Function", "State", "Detail")
    End If
    rowNumber = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(rowNumber, 1).Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(rowNumber, 2).Value2 = pCase
    If pCase > 0 Then
        ws.Cells(rowNumber, 3).Value2 = SuiteSheet("FunctionResults").Cells(pCase + 1, 1).Value2
        ws.Cells(rowNumber, 4).Value2 = SuiteSheet("FunctionResults").Cells(pCase + 1, 2).Value2
    End If
    ws.Cells(rowNumber, 5).Value2 = pState
    ws.Cells(rowNumber, 6).Value2 = Left$(pDetail, 30000)
End Sub

Private Sub RequireSuiteOwner()
    If ThisWorkbook.IsAddin Then Err.Raise 5, , "Run the suite in its disposable full-source XLSM"
    If ThisWorkbook.Path = "" Then Err.Raise 5, , "Save the disposable suite workbook before running"
    If LCase$(Right$(ThisWorkbook.Name, 5)) <> ".xlsm" Then Err.Raise 5, , "The suite owner must be an XLSM"
End Sub

Private Function SuiteVersionToken(ByVal pVersionText As String) As String
    Const prefix As String = "Stock Market Functions add-in, Version "
    Dim pathAt As Long
    If Left$(pVersionText, Len(prefix)) <> prefix Then Exit Function
    pathAt = InStr(1, pVersionText, " (", vbBinaryCompare)
    If pathAt <= Len(prefix) + 1 Then Exit Function
    SuiteVersionToken = Mid$(pVersionText, Len(prefix) + 1, pathAt - Len(prefix) - 1)
End Function

Private Function CheckSuiteVersion() As Boolean
    Dim control As Worksheet, actual As Variant, failure As String
    On Error GoTo Failed
    RequireSuiteOwner
    Set control = SuiteSheet("TestControl")
    control.Range("A4").Value2 = "Run ID": control.Range("B4").Value2 = SuiteRunID
    control.Range("A5").Value2 = "Expected SMF version": control.Range("B5").Value2 = ExpectedSMFVersion
    actual = RCHGetElementNumber("Version")
    control.Range("A6").Value2 = "SMF version": control.Range("B6").Value2 = CStr(actual)
    If VarType(actual) <> vbString Then Err.Raise 5, , "SMF version did not return text"
    If SuiteVersionToken(CStr(actual)) <> ExpectedSMFVersion Then Err.Raise 5, , "Unexpected SMF version"
    CheckSuiteVersion = True
    Exit Function
Failed:
    failure = CStr(Err.Number) & ":" & Err.Description
    On Error Resume Next
    If control Is Nothing Then Set control = SuiteSheet("TestControl")
    control.Range("B3").Value2 = "BLOCKED_VERSION"
    SuiteJournal 0, "BLOCKED_VERSION", failure
    ThisWorkbook.Save
    RestoreStatus
End Function

Public Sub SMFTestStop()
    RequireSuiteOwner
    CancelScheduled
    SuiteSheet("TestControl").Range("B3").Value2 = "STOP"
    SuiteJournal 0, "STOP", "Explicit stop; no case is retried"
    ThisWorkbook.Save
    RestoreStatus
End Sub

Public Sub SMFTestResume()
    Dim control As Worksheet, n As Long
    RequireSuiteOwner
    Set control = SuiteSheet("TestControl")
    If CStr(control.Range("B4").Value2) <> SuiteRunID Then Err.Raise 5, , "Run identity differs from this generated module"
    If Not CheckSuiteVersion() Then Exit Sub
    n = CLng(control.Range("B1").Value2)
    If n >= 1 And n <= SuiteCount Then
        If CStr(SuiteSheet("FunctionResults").Cells(n + 1, 6).Value2) = "RUNNING" Then
            control.Range("B3").Value2 = "BLOCKED_INTERRUPTED"
            SuiteJournal n, "BLOCKED_INTERRUPTED", "Resume refused: investigate the unfinished case before any retry"
            ThisWorkbook.Save
            RestoreStatus
            Exit Sub
        End If
    End If
    CaptureStatus
    control.Range("B3").Value2 = "RUN"
    SuiteJournal 0, "RESUME", "Resume from the completed-case cursor"
    ThisWorkbook.Save
    ScheduleNext
End Sub
'''

MODULE_STATE = '''Private mSavedStatus As Variant
Private mStatusCaptured As Boolean
Private mExecuting As Boolean
Private mScheduled As Boolean
Private mNextRun As Date

'''

START = '''
Public Sub SMFRunFullSuite()
    SMFTestStart
End Sub

Public Sub SMFTestStart(Optional ByVal firstCase As Long = 1, Optional ByVal lastCase As Long = 0)
    Dim ws As Worksheet, inputs As Worksheet, control As Worksheet, i As Long
    If mExecuting Then Err.Raise 5, , "A suite case is still executing"
    RequireSuiteOwner
    If Not CheckSuiteVersion() Then Exit Sub
    If lastCase = 0 Then lastCase = SuiteCount
    If firstCase < 1 Or lastCase < firstCase Or lastCase > SuiteCount Then Err.Raise 5, , "Invalid suite case range"
    CancelScheduled
    CaptureStatus
    Set ws = SuiteSheet("FunctionResults")
    Set inputs = SuiteSheet("TestInputs")
    Set control = SuiteSheet("TestControl")
    ws.Cells.ClearContents
    ws.Range("A1:Q" & CStr(SuiteCount + 1)).NumberFormat = "@"
    ws.Range("A1:Q1").Value = Array("Case", "Function", "Group", "Description", "Expected", "Status", "Seconds", "Actual", "Error", "Last HTTP status", "Last HTTP URL", "Last HTTP error", "Finished", "HTTP reset marker", "HTTP case ID", "HTTP run ID", "HTTP observation")
    For i = 1 To 5: inputs.Cells(i, 1).Value2 = i: Next i
    inputs.Range("J20").ColumnWidth = 30: inputs.Range("J20").RowHeight = 60
    control.Range("A1:A3").Value = Application.Transpose(Array("Next case", "Last case", "State"))
    control.Range("B1").Value2 = firstCase: control.Range("B2").Value2 = lastCase
    control.Range("B3").Value2 = "RUN"
    control.Range("A7").Value2 = "Started": control.Range("B7").Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    control.Range("A8").Value2 = "Owner workbook": control.Range("B8").Value2 = ThisWorkbook.FullName
    control.Range("A9").Value2 = "Case count": control.Range("B9").Value2 = SuiteCount
    control.Range("A10").Value2 = "Source fingerprint": control.Range("B10").Value2 = SuiteSourceFingerprint
    control.Range("A11").Value2 = "Case specification SHA256": control.Range("B11").Value2 = SuiteCasesSHA
    control.Range("A12").Value2 = "Build UTC": control.Range("B12").Value2 = SuiteBuildUTC
    control.Range("A13").Value2 = "First case": control.Range("B13").Value2 = firstCase
'''

RUN_HEAD = '''
Public Sub SMFTestRunNext()
    Dim ws As Worksheet, control As Worksheet, n As Long, row As Long, started As Double
    Dim result As Variant, obj As Object, ok As Boolean, detail As String
    Dim httpMarker As String, ignoredHTTP As String, httpURL As String, failure As String
    On Error GoTo RunnerFailed
    If mExecuting Then Exit Sub
    mScheduled = False
    RequireSuiteOwner
    Set control = SuiteSheet("TestControl")
    If control.Range("B3").Value2 <> "RUN" Then Exit Sub
    If CStr(control.Range("B4").Value2) <> SuiteRunID Then Err.Raise 5, , "Generated module/run identity mismatch"
    If Not CheckSuiteVersion() Then Exit Sub
    n = CLng(control.Range("B1").Value2)
    If n > CLng(control.Range("B2").Value2) Or n > SuiteCount Then
        control.Range("B3").Value2 = "DONE": ThisWorkbook.Save: RestoreStatus: Exit Sub
    End If
    If n < 1 Then Err.Raise 5, , "Invalid next-case cursor"
    Set ws = SuiteSheet("FunctionResults"): row = n + 1
    If CStr(ws.Cells(row, 6).Value2) = "RUNNING" Then
        control.Range("B3").Value2 = "BLOCKED_INTERRUPTED"
        SuiteJournal n, "BLOCKED_INTERRUPTED", "Previous case remains RUNNING; no automatic retry"
        ThisWorkbook.Save: RestoreStatus: Exit Sub
    End If
    mExecuting = True
    httpMarker = "smf-test-case:" & SuiteRunID & ":" & CStr(n)
    ignoredHTTP = smfMacFetch(httpMarker)
    ws.Range(ws.Cells(row, 7), ws.Cells(row, 17)).ClearContents
    ws.Cells(row, 14).Value2 = httpMarker
    ws.Cells(row, 6).Value2 = "RUNNING"
    SuiteJournal n, "RUNNING", "Saved before SETUP/CALL/CHECK"
    ThisWorkbook.Save
    started = Timer
    Application.StatusBar = "SMF test " & CStr(n) & "/" & CStr(SuiteCount) & ": " & ws.Cells(row, 1).Value2
    On Error GoTo CaseFailed
    Select Case n
'''

RUN_TAIL = '''
    End Select
    If ok Then ws.Cells(row, 6).Value2 = "PASS" Else ws.Cells(row, 6).Value2 = "CHECK_FAILED"
    GoTo Finished
CaseFailed:
    ws.Cells(row, 6).Value2 = "RUNTIME_ERROR"
    If InStr(Err.Source, "[CALL]") > 0 Then ws.Cells(row, 6).Value2 = "CALL_ERROR"
    If InStr(Err.Source, "[CHECK]") > 0 Then ws.Cells(row, 6).Value2 = "ASSERTION_ERROR"
    If InStr(Err.Source, "[SETUP]") > 0 Then ws.Cells(row, 6).Value2 = "SETUP_ERROR"
    If InStr(Err.Source, "[CLEANUP]") > 0 Then ws.Cells(row, 6).Value2 = "CLEANUP_ERROR"
    ws.Cells(row, 9).Value2 = CStr(Err.Number) & ": " & Err.Source & ": " & Err.Description
    Resume Finished
Finished:
    On Error GoTo RunnerFailed
    ws.Cells(row, 7).Value2 = CStr((Timer - started + 86400#) - 86400# * Int((Timer - started + 86400#) / 86400#))
    ws.Cells(row, 8).Value2 = Left$(SummarizeValue(result, obj) & IIf(Len(detail) > 0, "; " & detail, ""), 2000)
    ws.Cells(row, 10).Value2 = CStr(SMFMacHTTPStatus())
    httpURL = SMFMacLastURL()
    ws.Cells(row, 11).Value2 = httpURL
    ws.Cells(row, 12).Value2 = SMFMacLastError()
    ws.Cells(row, 13).Value2 = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    If httpURL <> httpMarker And (LCase$(Left$(httpURL, 7)) = "http://" Or LCase$(Left$(httpURL, 8)) = "https://") Then
        ws.Cells(row, 15).Value2 = ws.Cells(row, 1).Value2
        ws.Cells(row, 16).Value2 = SuiteRunID
        ws.Cells(row, 17).Value2 = "REQUEST"
    ElseIf httpURL = httpMarker Then
        ws.Cells(row, 17).Value2 = "NONE"
    Else
        ws.Cells(row, 17).Value2 = "OTHER"
    End If
    control.Range("B1").Value2 = n + 1
    SuiteJournal n, CStr(ws.Cells(row, 6).Value2), CStr(ws.Cells(row, 9).Value2)
    If n >= CLng(control.Range("B2").Value2) Then control.Range("B3").Value2 = "DONE"
    ThisWorkbook.Save
    mExecuting = False
    If control.Range("B3").Value2 = "RUN" Then ScheduleNext Else RestoreStatus
    Exit Sub
RunnerFailed:
    failure = CStr(Err.Number) & ": " & Err.Source & ": " & Err.Description
    On Error Resume Next
    mExecuting = False
    CancelScheduled
    If control Is Nothing Then Set control = SuiteSheet("TestControl")
    control.Range("B3").Value2 = "BLOCKED_RUNNER"
    SuiteJournal n, "BLOCKED_RUNNER", failure
    ThisWorkbook.Save
    RestoreStatus
End Sub
'''


def render_module(cases, metadata):
    out = ['Attribute VB_Name = "modSMFFunctionSuite"', 'Option Explicit',
           f'Private Const SuiteCount As Long = {len(cases)}',
           'Private Const SuiteRunID As String = ' + quoted(metadata["run_id"]),
           'Private Const ExpectedSMFVersion As String = ' + quoted(metadata["expected_smf_version"]),
           'Private Const SuiteSourceFingerprint As String = ' + quoted(metadata["source_fingerprint"]),
           'Private Const SuiteCasesSHA As String = ' + quoted(metadata["case_specification_sha256"]),
           'Private Const SuiteBuildUTC As String = ' + quoted(metadata["generated_utc"]),
           MODULE_STATE, HELPERS, START]
    metadata_subs = []
    for offset in range(0, len(cases), 25):
        sub_name = "InitializeCases" + str(offset // 25 + 1)
        out.append("    " + sub_name + " ws")
        lines = ["Private Sub " + sub_name + "(ByVal ws As Worksheet)"]
        for i, case in enumerate(cases[offset:offset + 25], offset + 1):
            for col, field in enumerate(("id", "function", "group", "description", "expected_text"), 1):
                lines.append(f"    ws.Cells({i + 1}, {col}).Value2 = {quoted(case[field])}")
            lines.append(f'    ws.Cells({i + 1}, 6).Value2 = "PENDING"')
        metadata_subs += lines + ["End Sub"]
    out += ['    ws.Columns("A:C").AutoFit',
            '    SuiteJournal 0, "START", "Fresh case statuses; expected version verified"',
            "    ThisWorkbook.Save", "    ScheduleNext", "End Sub", RUN_HEAD]
    for i in range(1, len(cases) + 1):
        out.append(f"        Case {i}: SuiteCase{i:03d} result, obj, ok, detail")
    out += [RUN_TAIL, *metadata_subs]
    for i, case in enumerate(cases, 1):
        out += [f"Private Sub SuiteCase{i:03d}(ByRef result As Variant, ByRef obj As Object, ByRef ok As Boolean, ByRef detail As String)",
                "    Dim savedNumber As Long, savedText As String, phase As String",
                "    Dim cleanupNumber As Long, cleanupText As String",
                *["    " + x for x in case.get("declarations", [])],
                '    phase = "SETUP"', "    On Error GoTo Failed",
                *["    " + x for x in case.get("setup", [])], '    phase = "CALL"',
                *["    " + x for x in case.get("call_steps", [])]]
        out.append("    " + ("Set obj = " if case.get("object_result") else "result = ") + case["expression"])
        out += ['    phase = "CHECK"', *["    " + x for x in case.get("check_setup", [])]]
        if case.get("result_details"):
            out.append("    detail = " + case["result_details"])
        # VBA And/Or do not short circuit. Preserve reviewed checks verbatim:
        # conversion failures are ASSERTION_ERROR, not provider/call failures.
        out += ["    ok = " + case["check"], "    GoTo CleanUp", "Failed:",
                "    savedNumber = Err.Number", "    savedText = Err.Description", "    Resume CleanUp",
                "CleanUp:", "    On Error Resume Next", "    Err.Clear",
                *["    " + x for x in case.get("cleanup", [])],
                "    cleanupNumber = Err.Number", "    cleanupText = Err.Description", "    On Error GoTo 0",
                f'    If savedNumber <> 0 Then Err.Raise savedNumber, {quoted(case["id"])} & " [" & phase & "]", savedText',
                f'    If cleanupNumber <> 0 Then Err.Raise cleanupNumber, {quoted(case["id"])} & " [CLEANUP]", cleanupText',
                "End Sub"]
    source = "\r\n".join(out) + "\r\n"
    if max(map(len, source.splitlines())) > 1023:
        raise ValueError("VBA physical line exceeds 1023 characters")
    if re.search(r"\{\{[A-Z_]+\}\}", source):
        raise ValueError("Unresolved fixture placeholder")
    return source


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--suite", type=Path, required=True, help="Disposable run output directory; created if absent.")
    endpoint = parser.add_mutually_exclusive_group(required=True)
    endpoint.add_argument("--server-config", type=Path, help="Explicit JSON containing a loopback base_url.")
    endpoint.add_argument("--local-http-base", help="Explicit loopback fixture URL; no configuration/state file is read.")
    parser.add_argument("--case-file", type=Path, help="Generate only selected reviewed cases; omit full-function requirement.")
    parser.add_argument("--expected-version", default=EXPECTED_VERSION)
    parser.add_argument("--run-id", help="Optional run identity shared with a fixture preparer.")
    args = parser.parse_args()
    try:
        suite = args.suite.expanduser().resolve()
        if any(c in str(suite) for c in '\r\n"'):
            raise ValueError('Suite paths cannot contain line breaks or double quotes in VBA fixture literals')
        if not args.expected_version.strip() or any(c in args.expected_version for c in '\r\n"'):
            raise ValueError("Expected version must be nonempty single-line text")
        if args.run_id and not re.fullmatch(r"[A-Za-z0-9_.-]{1,100}", args.run_id):
            raise ValueError("Run ID must contain 1..100 ASCII letters, digits, underscores, dots or dashes")
        config_path = args.server_config.expanduser().resolve() if args.server_config else None
        config_payload = config_path.read_bytes() if config_path else None
        config = json.loads(config_payload) if config_payload is not None else {}
        base_url = local_base(config["base_url"] if config_path else args.local_http_base)
        case_file = args.case_file.expanduser().resolve() if args.case_file else None
        cases, inputs = load_cases(case_file)
        cases = replace_tokens(cases, {"{{SUITE_DIR}}": str(suite), "{{LOCAL_HTTP_BASE}}": base_url,
                                      "{{SUITE_SHEET}}": "TestInputs"})
        functions = inventory()
        public = {f["function"].lower() for f in functions if f["visibility"].lower() == "public"}
        validate_cases(cases, public, not bool(case_file))
        case_payload = (json.dumps(cases, indent=2) + "\n").encode("utf-8")
        source_files = {}
        for path in sorted((REPO / "code").iterdir()):
            if path.suffix.lower() in {".bas", ".cls"} and path.name != "exportMods.bas":
                source_files[path.relative_to(REPO).as_posix()] = digest(path.read_bytes())
        fingerprint = digest(json.dumps(source_files, sort_keys=True).encode())
        now = datetime.now(timezone.utc)
        unsafe = [c["id"] for c in cases if re.search(r"\b(?:And|Or)\b", c["check"], re.I)
                  and re.search(r"\b(?:CDbl|CLng|CInt|CDate|CStr)\s*\(", c["check"], re.I)]
        metadata = {
            "schema_version": 1,
            "run_id": args.run_id or "smf-0.6-" + now.strftime("%Y%m%dT%H%M%SZ") + "-" + uuid4().hex[:10],
            "generated_utc": now.isoformat(), "source_root": str(REPO), "suite_dir": str(suite),
            "expected_smf_version": args.expected_version, "local_http_base": base_url,
            "server_config": str(config_path) if config_path else None,
            "server_config_sha256": digest(config_payload) if config_payload is not None else None,
            "case_inputs": inputs, "case_count": len(cases), "public_function_count": len(public),
            "covered_public_function_count": len(public & {c["function"].lower() for c in cases}),
            "full_function_coverage_required": not bool(case_file),
            "source_files_sha256": source_files, "source_fingerprint": fingerprint,
            "case_specification_sha256": digest(case_payload),
            "generator_sha256": digest(Path(__file__).read_bytes()),
            "worksheet_history_case_count": sum(c["id"].startswith("HistoryWorksheet_") for c in cases),
            "assertions_with_nonshortcircuit_conversions": unsafe,
            "assertion_policy": "Checks stay unchanged. Conversion errors during CHECK are ASSERTION_ERROR; CALL/SETUP/CLEANUP remain separate.",
            "http_provenance": "Each case resets the URL to its run/case marker without HTTP; O/P identify only subsequent HTTP(S) URLs.",
            "native_execution": "not_run",
            "coverage_scope": "Public Functions in .bas files; excludes Subs, Dictionary properties and arbitrary caller URLs.",
        }
        preparation = suite / "preparation.json"
        if preparation.exists():
            metadata["preparation_sha256"] = digest(preparation.read_bytes())
        source = render_module(cases, metadata)
        metadata["generated_module_sha256"] = digest(source.encode("utf-8"))
        suite.mkdir(parents=True, exist_ok=True)
        (suite / "modSMFFunctionSuite.bas").write_bytes(source.encode("utf-8"))
        (suite / "test-cases.json").write_bytes(case_payload)
        (suite / "function-inventory.json").write_text(json.dumps(functions, indent=2) + "\n")
        (suite / "run-manifest.json").write_text(json.dumps(metadata, indent=2) + "\n")
    except (OSError, ValueError, KeyError, TypeError) as error:
        parser.error(str(error))
    print(f"{len(public)} public functions in source; {len(cases)} selected cases; full function coverage required: {not bool(case_file)}")
    print(suite / "modSMFFunctionSuite.bas")
    print("SOURCE GENERATION ONLY: no Excel, VBA import, native calculation or network requests executed.")


if __name__ == "__main__":
    main()

Attribute VB_Name = "modMacJSONTests"
Option Explicit

Private smfMacJSONChecks As Long

' Run in the Immediate window or enter =SMFMacJSONSelfTest() in a cell.
' These fixtures use no website, cookie, account, or local file.
Public Function SMFMacJSONSelfTest() As String
#If Mac Then
    Dim fixture As String, rows As Variant, keys As Variant, path As Variant
    Dim node As Object, child As Object, deepJSON As String, deepPath As String
    Dim i As Long, parseError As Long

    smfMacJSONChecks = 0
    fixture = "{""a"":[{""raw"":0},{""raw"":false}],""keys"":{""0"":""zero"",""01"":""leading"",""1e2"":""exponent""}," & _
              """blank"":"""",""nil"":null,""negative"":-12.5,""id"":1234567890123456789}"

    smfMacJSONAssertNumber smfJSONExtractField("{""x"":7}", "x"), 7, "One field"
    smfMacJSONAssertNumber smfJSONExtractField("{""a"":{""b"":{""c"":7}}}", "a.b.c"), 7, "Three fields"
    smfMacJSONAssertNumber smfJSONExtractField("[{""raw"":7}]", "0.raw"), 7, "Root array"
    smfMacJSONAssertNumber smfJSONExtractField(" " & vbCrLf & "{""x"":7}" & vbTab & " ", "x"), 7, "JSON whitespace"
    smfMacJSONAssertNumber smfJSONExtractField(fixture, "a.0.raw"), 0, "Zero"
    smfMacJSONAssertNumber smfJSONExtractField_x64(fixture, "a.0.raw"), 0, "x64 entry point on Mac"
    smfMacJSONAssertBoolean smfJSONExtractField(fixture, "a.1.raw"), False, "False"
    smfMacJSONAssertText smfJSONExtractField(fixture, "keys.0"), "zero", "Numeric object key"
    smfMacJSONAssertText smfJSONExtractField(fixture, "keys.01"), "leading", "Leading-zero object key"
    smfMacJSONAssertText smfJSONExtractField(fixture, "keys.1e2"), "exponent", "Exponent-like object key"
    smfMacJSONAssertText smfJSONExtractField(fixture, "blank"), "", "Empty string"
    smfMacJSONAssert IsNull(smfJSONExtractField(fixture, "nil")), "JSON null"
    smfMacJSONAssertNumber smfJSONExtractField(fixture, "negative"), -12.5, "Negative decimal"
    smfMacJSONAssertText smfJSONExtractField(fixture, "id"), "1234567890123456789", "Large number stays text"
    smfMacJSONAssertText smfJSONExtractField(fixture, "missing"), "Not Found", "Absent field"
    smfMacJSONAssertText smfJSONExtractField(fixture, "negative.child"), "Not Found", "Scalar intermediate node"
    smfMacJSONAssertText smfJSONExtractField(fixture, "nil.child"), "Not Found", "Null intermediate node"

    For Each path In Array("-1", "1.5", "1e2", "01", "99", String$(400, "9"))
        smfMacJSONAssertText smfJSONExtractField("[1]", CStr(path)), "Not Found", "Invalid array index"
    Next path

    For i = 1 To 11
        deepJSON = deepJSON & "{""k"":"
        If i > 1 Then deepPath = deepPath & "."
        deepPath = deepPath & "k"
    Next i
    deepJSON = deepJSON & "7" & String$(11, "}")
    smfMacJSONAssertNumber smfJSONExtractField(deepJSON, deepPath), 7, "More than ten fields"

    InitScriptEngine
    Set node = DecodeJsonString(fixture)
    Set child = GetObjectProperty(node, "keys")
    smfMacJSONAssertText GetProperty(child, "01"), "leading", "Object helper"
    smfMacJSONAssert IsEmpty(GetProperty(child, "missing")), "Missing property helper"
    keys = GetKeys(child)
    smfMacJSONAssert LBound(keys) = 0 And UBound(keys) = 2, "Object keys array"
    smfMacJSONAssertText keys(0), "0", "First object key"

    rows = smfJSONExtractKeys(fixture, "keys")
    smfMacJSONAssert LBound(rows, 1) = 1 And UBound(rows, 1) = 4, "Key output rows"
    smfMacJSONAssert LBound(rows, 2) = 1 And UBound(rows, 2) = 1, "Key output column"
    smfMacJSONAssertText rows(1, 1), "0", "Numeric key output"
    smfMacJSONAssertText rows(2, 1), "01", "Leading-zero key output"
    smfMacJSONAssertText rows(4, 1), "#N/A", "Final key sentinel"

    rows = smfJSONExtractKeys(fixture, "a")
    smfMacJSONAssertText rows(1, 1), "0", "Array first index"
    smfMacJSONAssertText rows(2, 1), "1", "Array second index"
    smfMacJSONAssertText rows(3, 1), "#N/A", "Array key sentinel"
    rows = smfJSONExtractKeys("{}", "")
    smfMacJSONAssert UBound(rows, 1) = 1, "Empty object"
    smfMacJSONAssertText rows(1, 1), "#N/A", "Empty object sentinel"
    rows = smfJSONExtractKeys("[]", "")
    smfMacJSONAssert UBound(rows, 1) = 1, "Empty array"
    smfMacJSONAssertText rows(1, 1), "#N/A", "Empty array sentinel"
    smfMacJSONAssertKeysFail fixture, "missing"
    smfMacJSONAssertKeysFail fixture, "negative"
    smfMacJSONAssertKeysFail fixture, "nil"

    Set node = JsonConverter.ParseJson(fixture)
    Set node = JsonConverter.ParseJson(JsonConverter.ConvertToJson(node))
    Set child = GetObjectProperty(node, "keys")
    smfMacJSONAssertText GetProperty(child, "0"), "zero", "Dictionary JSON round trip"

    For Each path In Array("[", "[1", "{""a"":", "{}garbage", "[]{}", "{""a"":1}<html>error</html>")
        On Error Resume Next
        Set node = JsonConverter.ParseJson(CStr(path))
        parseError = Err.Number
        Err.Clear
        On Error GoTo 0
        smfMacJSONAssert parseError = 10001, "Invalid JSON reports parse error"
    Next path

    SMFMacJSONSelfTest = "PASS: " & smfMacJSONChecks & " JSON checks"
#Else
    Err.Raise 5, "SMFMacJSONSelfTest", "Run this test in Excel for Mac"
#End If
End Function

Private Sub smfMacJSONAssert(ByVal condition As Boolean, ByVal label As String)
    If Not condition Then Err.Raise vbObjectError + 2101, "SMFMacJSONSelfTest", label
    smfMacJSONChecks = smfMacJSONChecks + 1
End Sub

Private Sub smfMacJSONAssertNumber(ByVal actual As Variant, ByVal expected As Double, ByVal label As String)
    smfMacJSONAssert IsNumeric(actual) And VarType(actual) <> vbString, label & " type"
    smfMacJSONAssert CDbl(actual) = expected, label & " value"
End Sub

Private Sub smfMacJSONAssertText(ByVal actual As Variant, ByVal expected As String, ByVal label As String)
    smfMacJSONAssert VarType(actual) = vbString, label & " type"
    smfMacJSONAssert CStr(actual) = expected, label & " value"
End Sub

Private Sub smfMacJSONAssertBoolean(ByVal actual As Variant, ByVal expected As Boolean, ByVal label As String)
    smfMacJSONAssert VarType(actual) = vbBoolean, label & " type"
    smfMacJSONAssert CBool(actual) = expected, label & " value"
End Sub

Private Sub smfMacJSONAssertKeysFail(ByVal fixture As String, ByVal path As String)
    Dim result As Variant, errorNumber As Long
    On Error Resume Next
    result = smfJSONExtractKeys(fixture, path)
    errorNumber = Err.Number
    Err.Clear
    On Error GoTo 0
    smfMacJSONAssert errorNumber <> 0, "Invalid keys path fails: " & path
End Sub

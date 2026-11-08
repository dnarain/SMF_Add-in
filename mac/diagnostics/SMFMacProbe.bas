Attribute VB_Name = "SMFMacProbe"
Option Explicit

' Import only into a dedicated disposable diagnostic workbook.
' No HTTP, AppleScriptTask, VBProject access, or other SMF module is used.
Public Function SMFMacProbeBasic() As String
    Dim value As Long
    Dim parts As Variant
    value = 21 + 21
    If value <> 42 Then Err.Raise 5, , "Arithmetic check failed"
    If Replace("A" & vbLf & "B", vbLf, "|") <> "A|B" Then
        Err.Raise 5, , "String check failed"
    End If
    parts = Split("200|body", "|")
    If UBound(parts) <> 1 Or parts(1) <> "body" Then
        Err.Raise 5, , "Array check failed"
    End If
    SMFMacProbeBasic = "PASS: basic VBA (3 checks)"
End Function

Public Sub SMFMacProbeBasicToSheet()
    ThisWorkbook.Worksheets(1).Range("A1").Value = SMFMacProbeBasic()
End Sub

Attribute VB_Name = "finish_SMFMac"
Option Explicit

' Import into the disposable project after all code/ modules.
' This finish step uses no VBProject access and changes no security setting.
Public Sub SMFMacFinishInteractive()
    Dim projectFolder As Variant, result As String
    projectFolder = Application.InputBox("Enter the full POSIX path of the SMF project folder.", _
                                        "Finish SMF Mac build", ThisWorkbook.Path, Type:=2)
    If VarType(projectFolder) = vbBoolean Then Exit Sub
    result = SMFMacFinishBuild(CStr(projectFolder))
    If Left$(result, 6) = "FAILED" Then
        MsgBox result, vbExclamation, "SMF Mac build"
    Else
        MsgBox result, vbInformation, "SMF Mac build"
    End If
End Sub

Public Function SMFMacFinishBuild(ByVal projectFolder As String) As String
    Dim stage As String, filename As String, saveFilename As String, owner As String
    Dim transport As Variant, html As Variant, json As Variant, integration As Variant
    On Error GoTo Failed
    owner = ThisWorkbook.Name
    If Left$(projectFolder, 1) <> "/" Or Len(projectFolder) < 2 Then
        SMFMacFinishBuild = "FAILED: Use an absolute POSIX project path"
        Exit Function
    End If
    Do While Right$(projectFolder, 1) = "/"
        projectFolder = Left$(projectFolder, Len(projectFolder) - 1)
    Loop
    filename = projectFolder & "/mac/out/SMF-Mac.xlam"
    stage = "Check the target path"
    If Len(Dir$(filename)) > 0 Then
        SMFMacFinishBuild = "FAILED: An add-in already exists at " & filename
        Exit Function
    End If
    stage = "Run the transport fixtures"
    transport = Application.Run("'" & owner & "'!SMFMacTransportSelfTest")
    stage = "Run the HTML fixtures"
    html = Application.Run("'" & owner & "'!smfMacHTMLSelfTest")
    stage = "Run the JSON fixtures"
    json = Application.Run("'" & owner & "'!SMFMacJSONSelfTest")
    stage = "Run the integration fixtures"
    integration = Application.Run("'" & owner & "'!SMFMacIntegrationSelfTest")
    stage = "Save the Mac add-in"
    saveFilename = filename
#If Mac Then
    saveFilename = AppleScriptTask("SMFMac.scpt", "excelPath", filename)
#End If
    ThisWorkbook.IsAddin = True
    ThisWorkbook.SaveAs Filename:=saveFilename, FileFormat:=55
    SMFMacFinishBuild = "Saved: " & ThisWorkbook.FullName & vbCrLf & _
        "Transport fixtures: " & CStr(transport) & vbCrLf & _
        "HTML fixtures: " & CStr(html) & vbCrLf & _
        "JSON fixtures: " & CStr(json) & vbCrLf & _
        "Integration fixtures: " & CStr(integration)
    Exit Function
Failed:
    SMFMacFinishBuild = "FAILED: " & stage & ": " & CStr(Err.Number) & ": " & Err.Description
End Function

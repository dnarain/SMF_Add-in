Attribute VB_Name = "build_SMFMac_noprompt"
Option Explicit

' Experimental alternative builder. Native execution remains unverified.
' The helper stages source bytes inside Excel's own cache before each import.
' The default leaves a new, unsaved workbook open for inspection.
' Optional output can still require permission to write outside the container.
' This module never changes macro security or VBA project access.

' Diagnostic entry point: return without an explicit target test or compile call.
Public Function SMFMacPrepareCompile(ByVal projectFolder As String) As String
    SMFMacPrepareCompile = SMFMacBuildNoPrompt(projectFolder, False, False)
End Function

Public Function SMFMacBuildNoPrompt(ByVal projectFolder As String, _
                                   Optional ByVal runTests As Boolean = True, _
                                   Optional ByVal saveOutput As Boolean = False) As String
#If Mac Then
    Const xlOpenXMLAddIn As Long = 55
    Const xlWBATWorksheet As Long = -4167
    Dim target As Workbook, project As Object, component As Object
    Dim filenames As Variant, filename As Variant, stagedPath As String
    Dim targetName As String, targetFullName As String, outputFile As String
    Dim targetProjectName As String, componentCount As Long
    Dim expectedType As Long, actualType As Long
    Dim excelSavePath As String, imported As Long, targetSaved As Boolean
    Dim transportResult As Variant, htmlResult As Variant
    Dim jsonResult As Variant, integrationResult As Variant
    Dim oldAlerts As Boolean, oldUpdating As Boolean, oldEvents As Boolean
    Dim flagsRestored As Boolean
    Dim stage As String, errorNumber As Long, errorText As String, result As String

    oldAlerts = Application.DisplayAlerts
    oldUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    On Error GoTo BuildFailed
    stage = "Check the project path"
    projectFolder = Trim$(projectFolder)
    Do While Len(projectFolder) > 1 And Right$(projectFolder, 1) = "/"
        projectFolder = Left$(projectFolder, Len(projectFolder) - 1)
    Loop
    If Left$(projectFolder, 1) <> "/" Or Len(projectFolder) = 1 Then
        Err.Raise vbObjectError + 2460, , "Use an absolute POSIX project-folder path."
    End If
    If InStr(projectFolder, vbCr) > 0 Or InStr(projectFolder, vbLf) > 0 Then
        Err.Raise vbObjectError + 2461, , "The project path cannot contain line breaks."
    End If
    filenames = smfMacNoPromptSourceNames()
    smfMacNoPromptLog projectFolder, "Begin build; owner=" & ThisWorkbook.Name & _
        "; runTests=" & CStr(runTests) & "; saveOutput=" & CStr(saveOutput)

    stage = "Create a new build workbook"
    smfMacNoPromptLog projectFolder, stage
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set target = Application.Workbooks.Add(xlWBATWorksheet)
    targetName = target.Name

    stage = "Access the new workbook's VBA project"
    smfMacNoPromptLog projectFolder, stage & "; target=" & targetName
    Set project = target.VBProject
    targetProjectName = "SMFDiag_" & Format$(Now, "yyyymmdd_hhnnss")
    project.Name = targetProjectName
    target.Worksheets(1).Range("A1").Value = targetProjectName
    smfMacNoPromptLog projectFolder, "Target identity; workbook=" & targetName & _
        "; path=" & target.Path & "; project=" & project.Name
    imported = project.VBComponents.Count
    imported = 0

    For Each filename In filenames
        stage = "Stage module " & CStr(filename)
        smfMacNoPromptLog projectFolder, stage & "; imported=" & CStr(imported)
        stagedPath = AppleScriptTask("SMFMac.scpt", "stageModule", _
                                    projectFolder & vbLf & CStr(filename))
        smfMacNoPromptCheckStagedPath stagedPath, CStr(filename)
        stage = "Import module " & CStr(filename)
        smfMacNoPromptLog projectFolder, stage & "; imported=" & CStr(imported)
        Set component = project.VBComponents.Import(stagedPath)
        expectedType = 1 ' vbext_ct_StdModule
        If LCase$(Right$(CStr(filename), 4)) = ".cls" Then expectedType = 2 ' vbext_ct_ClassModule
        actualType = CLng(component.Type)
        If actualType <> expectedType Then
            Err.Raise vbObjectError + 2466, "SMFMacBuildNoPrompt", _
                "Imported component type mismatch for " & CStr(filename) & _
                ": expected type " & CStr(expectedType) & "; actual type " & CStr(actualType) & "."
        End If
        If StrComp(component.Name, Left$(CStr(filename), Len(CStr(filename)) - 4), vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 2462, , "The imported component name differs from the source filename."
        End If
        imported = imported + 1
    Next filename

    componentCount = project.VBComponents.Count
    smfMacNoPromptLog projectFolder, "Imports completed; target=" & targetName & _
        "; project=" & targetProjectName & "; imported=" & CStr(imported) & _
        "; components=" & CStr(componentCount)
    If runTests Then
        stage = "Run the transport fixtures"
        smfMacNoPromptLog projectFolder, stage
        transportResult = Application.Run("'" & targetName & "'!SMFMacTransportSelfTest")
        stage = "Run the HTML fixtures"
        smfMacNoPromptLog projectFolder, stage
        htmlResult = Application.Run("'" & targetName & "'!smfMacHTMLSelfTest")
        stage = "Run the JSON fixtures"
        smfMacNoPromptLog projectFolder, stage
        jsonResult = Application.Run("'" & targetName & "'!SMFMacJSONSelfTest")
        stage = "Run the integration fixtures"
        smfMacNoPromptLog projectFolder, stage
        integrationResult = Application.Run("'" & targetName & "'!SMFMacIntegrationSelfTest")
        smfMacNoPromptLog projectFolder, "Fixtures completed; target=" & targetName
    Else
        transportResult = "not run"
        htmlResult = "not run"
        jsonResult = "not run"
        integrationResult = "not run"
    End If

    If saveOutput Then
        stage = "Obtain the output path"
        smfMacNoPromptLog projectFolder, stage
        outputFile = projectFolder & "/mac/out/SMF-Mac.xlam"
        excelSavePath = AppleScriptTask("SMFMac.scpt", "excelPath", outputFile)
        stage = "Save the Mac add-in"
        smfMacNoPromptLog projectFolder, stage
        ' Keep overwrite and write-permission requests visible in this optional phase.
        Application.DisplayAlerts = True
        target.IsAddin = True
        target.SaveAs Filename:=excelSavePath, FileFormat:=xlOpenXMLAddIn
        targetName = target.Name
        targetFullName = target.FullName
        targetSaved = target.Saved
        result = "SaveAs completed. Verify the reported file separately."
    Else
        targetSaved = False
        result = "Built in memory. No add-in file was created."
    End If

    result = result & vbCrLf & "Target workbook: " & targetName & _
        vbCrLf & "Target path: " & target.Path & _
        vbCrLf & "Target VBA project: " & targetProjectName & _
        vbCrLf & "VBA components: " & CStr(componentCount) & _
        vbCrLf & "saved=" & CStr(targetSaved) & vbCrLf & _
        "Imported modules: " & CStr(imported) & vbCrLf & _
        "Transport fixtures: " & CStr(transportResult) & vbCrLf & _
        "HTML fixtures: " & CStr(htmlResult) & vbCrLf & _
        "JSON fixtures: " & CStr(jsonResult) & vbCrLf & _
        "Integration fixtures: " & CStr(integrationResult)
    If Len(targetFullName) > 0 Then result = result & vbCrLf & "Reported FullName: " & targetFullName
    result = result & vbCrLf & "The target remains open for inspection. Live sources were not tested."
    smfMacNoPromptRestoreFlags oldAlerts, oldUpdating, oldEvents
    flagsRestored = (Application.DisplayAlerts = oldAlerts) And _
        (Application.ScreenUpdating = oldUpdating) And (Application.EnableEvents = oldEvents)
    result = result & vbCrLf & "DisplayAlerts=" & CStr(Application.DisplayAlerts) & _
        "; ScreenUpdating=" & CStr(Application.ScreenUpdating) & _
        "; EnableEvents=" & CStr(Application.EnableEvents)
    result = result & vbCrLf & "FlagsRestored=" & CStr(flagsRestored)
    smfMacNoPromptLog projectFolder, "Completed after restoring flags; target=" & targetName & _
        "; project=" & targetProjectName & "; saved=" & CStr(targetSaved) & _
        "; runTests=" & CStr(runTests) & "; saveOutput=" & CStr(saveOutput) & _
        "; DisplayAlerts=" & CStr(Application.DisplayAlerts) & _
        "; ScreenUpdating=" & CStr(Application.ScreenUpdating) & _
        "; EnableEvents=" & CStr(Application.EnableEvents) & _
        "; FlagsRestored=" & CStr(flagsRestored)
    SMFMacBuildNoPrompt = result
    Exit Function

BuildFailed:
    errorNumber = Err.Number
    errorText = Err.Description
    smfMacNoPromptLog projectFolder, "FAILED; stage=" & stage & "; error=" & CStr(errorNumber) & _
        "; imported=" & CStr(imported) & "; description=" & errorText
    On Error Resume Next
    If Not target Is Nothing Then
        If Not targetSaved Then target.IsAddin = False
        targetName = target.Name
        targetSaved = target.Saved
    End If
    smfMacNoPromptRestoreFlags oldAlerts, oldUpdating, oldEvents
    On Error GoTo 0
    result = "FAILED: " & CStr(errorNumber) & vbCrLf & stage & ": " & errorText & _
        vbCrLf & "Imported modules: " & CStr(imported)
    If Len(targetName) > 0 Then
        result = result & vbCrLf & "Target workbook: " & targetName & _
            vbCrLf & "Target VBA project: " & targetProjectName & _
            vbCrLf & "saved=" & CStr(targetSaved) & vbCrLf & "The owned target remains open for inspection."
    Else
        result = result & vbCrLf & "No target workbook was created."
    End If
    SMFMacBuildNoPrompt = result
#Else
    SMFMacBuildNoPrompt = "FAILED: This builder requires native Excel for Mac."
#End If
End Function

Private Sub smfMacNoPromptLog(ByVal projectFolder As String, ByVal stateText As String)
#If Mac Then
    Dim ignored As String
    ' Logging is diagnostic. A logging error must not interrupt the build.
    On Error Resume Next
    stateText = Replace(Replace(stateText, vbCr, " "), vbLf, " ")
    ignored = AppleScriptTask("SMFMac.scpt", "logBuildState", projectFolder & vbLf & stateText)
#End If
End Sub

Private Sub smfMacNoPromptRestoreFlags(ByVal alerts As Boolean, ByVal updating As Boolean, ByVal events As Boolean)
    ' Restore each flag even if Excel rejects another restoration request.
    On Error Resume Next
    Application.DisplayAlerts = alerts
    Application.ScreenUpdating = updating
    Application.EnableEvents = events
End Sub

Private Sub smfMacNoPromptCheckStagedPath(ByVal stagedPath As String, ByVal filename As String)
    Const cacheFolder = ":Library:Containers:com.microsoft.Excel:Data:Library:Caches:SMF-Mac-Build:v0.2:"
    If InStr(stagedPath, vbCr) > 0 Or InStr(stagedPath, vbLf) > 0 Then
        Err.Raise vbObjectError + 2463, , "The staged module path contains a line break."
    End If
    If InStr(1, stagedPath, cacheFolder, vbBinaryCompare) = 0 Then
        Err.Raise vbObjectError + 2464, , "The helper did not return a path inside Excel's SMF build cache."
    End If
    If Right$(stagedPath, Len(filename) + 1) <> ":" & filename Then
        Err.Raise vbObjectError + 2465, , "The staged module filename differs from the requested filename."
    End If
End Sub

Private Function smfMacNoPromptSourceNames() As Variant
    Dim filenames As String
    ' BEGIN GENERATED SOURCE FILES
    filenames = "Dictionary.cls|JsonConverter.bas|modExtractData.bas|modGetAdvFNElement.bas|modGetCSVFile.bas"
    filenames = filenames & "|modGetEconData.bas|modGetElementNumber.bas|modGetGuruFocusData.bas|modGetGuruFocusItem.bas|modGetHTMLTable.bas"
    filenames = filenames & "|modGetMSNHistory.bas|modGetOptionExpirations.bas|modGetOptionExpiry.bas|modGetOptionQuotes.bas|modGetOptionStrikes.bas"
    filenames = filenames & "|modGetTableCell.bas|modGetTagContent.bas|modGetWebData.bas|modGetYahooHistory.bas|modGetYahooJSONData.bas"
    filenames = filenames & "|modGetYahooQuotes.bas|modInCellChart.bas|modJSONExtract.bas|modMacCode.bas|modMacHTML.bas"
    filenames = filenames & "|modMacIntegrationTests.bas|modMacJSONTests.bas|modMenu.bas|modPictureInsert.bas|modPictureTable.bas"
    filenames = filenames & "|modPictureToggle.bas|modRSITargetPrices.bas|modTechnicalAnalysis.bas|modYahooAPIData.bas|modZAdditional.bas"
    filenames = filenames & "|smfConvertData_.bas|smfConvertYahooTicker_.bas|smfCreateComment_.bas|smfDownloadTable_.bas|smfGetGoogleHistoryCSV_.bas"
    filenames = filenames & "|smfGetGuruFocusCSVItem_.bas|smfGetMorningStarCSV_.bas|smfGetYahooHistoryCSV_.bas|smfGetYahooHistory_.bas|smfGetYahooPortfolio_.bas"
    filenames = filenames & "|smfPricesBetween_.bas|smfPricesByDates_.bas|smfUtilities_.bas|smfxDatabaseUpdate_.bas"
    ' END GENERATED SOURCE FILES
    smfMacNoPromptSourceNames = Split(filenames, "|")
End Function

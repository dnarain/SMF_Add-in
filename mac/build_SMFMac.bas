Attribute VB_Name = "build_SMFMac"
Option Explicit

' Import this one module into a new, disposable Excel workbook.
' Run SMFMacBuildInteractive from Tools > Macro > Macros.
' The builder creates another workbook and imports code/ into that workbook.
' It never edits the workbook that contains this builder.
' The builder does not change Excel macro security or VBA project access.

Public Sub SMFMacBuildInteractive()
    Dim projectFolder As Variant, result As String
    projectFolder = Application.InputBox( _
        "Enter the full POSIX path of the smf-mac project folder.", _
        "Build SMF for Mac", ThisWorkbook.Path, Type:=2)
    If VarType(projectFolder) = vbBoolean Then Exit Sub
    On Error GoTo BuildFailed
    result = SMFMacBuild(CStr(projectFolder), True)
    If Left$(result, 7) = "FAILED:" Then
        MsgBox result, vbExclamation, "SMF Mac build failed"
    Else
        MsgBox result, vbInformation, "SMF Mac build"
    End If
    Exit Sub
BuildFailed:
    MsgBox Err.Description, vbExclamation, "SMF Mac build failed"
End Sub

Public Function SMFMacBuild(ByVal projectFolder As String, _
                           Optional ByVal runSelfTests As Boolean = True) As String
    Const xlOpenXMLAddIn As Long = 55
    Const xlWBATWorksheet As Long = -4167
    Dim sourceFolder As String, outputFolder As String, outputFile As String
    Dim excelSavePath As String, actualFullName As String
    Dim target As Workbook, project As Object, component As Object
    Dim modules As Collection, classes As Collection, sourceFile As Variant
    Dim sourcePaths As Variant
    Dim imported As Long, oldAlerts As Boolean, oldUpdating As Boolean
    Dim oldEvents As Boolean, stage As String, htmlResult As Variant, jsonResult As Variant
    Dim transportResult As Variant, integrationResult As Variant
    Dim saved As Boolean, errorNumber As Long, errorText As String

    oldAlerts = Application.DisplayAlerts
    oldUpdating = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    On Error GoTo BuildFailed
    stage = "Check the project path"
    projectFolder = Trim$(projectFolder)
    Do While Len(projectFolder) > 1 And Right$(projectFolder, 1) = "/"
        projectFolder = Left$(projectFolder, Len(projectFolder) - 1)
    Loop
    If Left$(projectFolder, 1) <> "/" Then
        Err.Raise vbObjectError + 2400, , "Enter an absolute POSIX path, such as /Users/name/Documents/smf-mac."
    End If
    sourceFolder = projectFolder & "/code/"
    outputFolder = projectFolder & "/mac/out"
    outputFile = outputFolder & "/SMF-Mac.xlam"

#If Mac Then
    stage = "Request access to project source files"
    sourcePaths = smfMacBuildSourcePaths(sourceFolder)
    If Not GrantAccessToMultipleFiles(sourcePaths) Then
        Err.Raise vbObjectError + 2404, , "Excel did not receive access to the project source files."
    End If
    ' The explicit manifest avoids a second permission for source-folder enumeration.
    Set classes = New Collection
    Set modules = New Collection
    For Each sourceFile In sourcePaths
        If LCase$(Right$(CStr(sourceFile), 4)) = ".cls" Then
            classes.Add CStr(sourceFile)
        Else
            modules.Add CStr(sourceFile)
        End If
    Next sourceFile
#Else
    Set classes = smfMacBuildFiles(sourceFolder, "*.cls")
    Set modules = smfMacBuildFiles(sourceFolder, "*.bas")
#End If

    stage = "Check source files and the output folder"
    If Len(Dir$(sourceFolder & "modGetElementNumber.bas")) = 0 Then
        Err.Raise vbObjectError + 2401, , "The project folder must contain code/modGetElementNumber.bas."
    End If
    If Len(Dir$(sourceFolder & "Dictionary.cls")) = 0 Then
        Err.Raise vbObjectError + 2402, , "The Mac dictionary is missing: code/Dictionary.cls."
    End If
    If Len(Dir$(outputFolder, vbDirectory)) = 0 Then MkDir outputFolder
    If Len(Dir$(outputFile)) > 0 Then
        Err.Raise vbObjectError + 2403, , "SMF-Mac.xlam already exists. Move that file before another build."
    End If
    stage = "Create a new add-in workbook"
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Set target = Application.Workbooks.Add(xlWBATWorksheet)

    stage = "Access the new workbook's VBA project"
    Set project = target.VBProject
    ' Reading VBComponents checks access before the builder imports source files.
    imported = project.VBComponents.Count
    imported = 0

    stage = "Import class modules"
    For Each sourceFile In classes
        Set component = project.VBComponents.Import(CStr(sourceFile))
        imported = imported + 1
    Next sourceFile

    stage = "Import standard modules"
    For Each sourceFile In modules
        ' The upstream developer export tool needs VBIDE references and a Windows drive.
        If LCase$(CStr(sourceFile)) <> LCase$(sourceFolder & "exportMods.bas") Then
            Set component = project.VBComponents.Import(CStr(sourceFile))
            imported = imported + 1
        End If
    Next sourceFile

    stage = "Run the transport, HTML, JSON, and integration self-tests"
    If runSelfTests Then
        transportResult = Application.Run("'" & target.Name & "'!SMFMacTransportSelfTest")
        htmlResult = Application.Run("'" & target.Name & "'!smfMacHTMLSelfTest")
        jsonResult = Application.Run("'" & target.Name & "'!SMFMacJSONSelfTest")
        integrationResult = Application.Run("'" & target.Name & "'!SMFMacIntegrationSelfTest")
    Else
        transportResult = "not run"
        htmlResult = "not run"
        jsonResult = "not run"
        integrationResult = "not run"
    End If

    stage = "Save the Mac add-in"
    target.IsAddin = True
    Application.DisplayAlerts = False
    excelSavePath = outputFile
#If Mac Then
    excelSavePath = AppleScriptTask("SMFMac.scpt", "excelPath", outputFile)
#End If
    target.SaveAs Filename:=excelSavePath, FileFormat:=xlOpenXMLAddIn
    actualFullName = target.FullName
    saved = True
    target.Close SaveChanges:=False
    Set target = Nothing
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldUpdating
    Application.EnableEvents = oldEvents
    SMFMacBuild = "Saved: " & actualFullName & vbCrLf & _
        "Imported modules: " & CStr(imported) & vbCrLf & _
        "Transport self-tests passed: " & CStr(transportResult) & vbCrLf & _
        "HTML self-tests passed: " & CStr(htmlResult) & vbCrLf & _
        "JSON self-tests: " & CStr(jsonResult) & vbCrLf & _
        "Integration self-tests passed: " & CStr(integrationResult) & vbCrLf & _
        "The self-tests use local fixtures. They do not verify live data sources."
    Exit Function

BuildFailed:
    errorNumber = Err.Number
    errorText = Err.Description
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not target Is Nothing Then target.Close SaveChanges:=False
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldUpdating
    Application.EnableEvents = oldEvents
    On Error GoTo 0
    If stage = "Access the new workbook's VBA project" Then
        errorText = errorText & vbCrLf & _
            "Excel did not allow VBA project access. The builder did not change that setting."
    End If
    If saved Then errorText = errorText & vbCrLf & "A file was saved before this error: " & outputFile
    ' Return caught errors to the caller. An unhandled VBA error opens a modal
    ' dialog and prevents an AppleScript caller from receiving the result.
    SMFMacBuild = "FAILED: " & CStr(errorNumber) & vbCrLf & stage & ": " & errorText
End Function

Private Function smfMacBuildSourcePaths(ByVal sourceFolder As String) As Variant
    Dim filenames As String, paths As Variant, i As Long
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
    paths = Split(filenames, "|")
    For i = LBound(paths) To UBound(paths)
        paths(i) = sourceFolder & CStr(paths(i))
    Next i
    smfMacBuildSourcePaths = paths
End Function

Private Function smfMacBuildFiles(ByVal folder As String, ByVal pattern As String) As Collection
    Dim result As New Collection, filename As String
    filename = Dir$(folder & pattern)
    Do While Len(filename) > 0
        result.Add folder & filename
        filename = Dir$()
    Loop
    Set smfMacBuildFiles = result
End Function

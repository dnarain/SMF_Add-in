Attribute VB_Name = "modMacIntegrationTests"
Option Explicit

Public Function SMFMacIntegrationSelfTest() As Long
    Dim html As String, actual As Variant, recipeChecks As Long
    html = "<table><tr><td>Earnings</td><td><b>May 02 AMC</b></td>" & _
           "<td>Price</td><td><b>123.456789</b></td></tr></table>"
    actual = smfGetTagContent(html, "TD", 1, ">Earnings", " ", " ", " ", 1, "Error", 0)
    If CStr(actual) <> "May 02 AMC" Then Err.Raise 5, , "Earnings HTML extraction failed"
    actual = smfGetTagContent(html, "TD", 1, ">Price", " ", " ", " ", 1, "Error", 0)
    If Not IsNumeric(actual) Then Err.Raise 5, , "Price HTML extraction is not numeric"
    If Abs(CDbl(actual) - 123.456789) > 0.000000001 Then Err.Raise 5, , "Price precision changed"
    actual = smfCDec("0.123456789")
    If Abs(CDbl(actual) - 0.123456789) > 0.000000000001 Then Err.Raise 5, , "Decimal conversion lost precision"
    actual = smfGetTagContent(html, "TD", 1, ">Missing", " ", " ", " ", 1, "Missing", 0)
    If CStr(actual) <> "Missing" Then Err.Raise 5, , "Missing field became a numeric value"
#If Mac Then
    If SMFMacElementsSelfTest() <> 3 Then Err.Raise 5, , "Element text fixture failed"
    recipeChecks = SMFMacRecipeSelfTest()
#End If
    SMFMacIntegrationSelfTest = 5 + recipeChecks
End Function

Public Function SMFMacRecipeSelfTest() As Long
    Dim html As String, quotedHTML As String, prefix As String, recipe As String
    Dim actual As Variant, rejected As Variant, item As Variant, checks As Long
    Dim beforeURL As String, beforeError As String, beforeStatus As Long
    beforeURL = SMFMacLastURL()
    beforeError = SMFMacLastError()
    beforeStatus = SMFMacHTTPStatus()
    html = "<table><tr><td>Earnings</td><td><b>May 02 AMC</b></td>" & _
           "<td>Price</td><td><b>123.456789</b></td></tr></table>"
    quotedHTML = smfMacFixtureQuote(html)
    prefix = "=smfGetTagContent(" & quotedHTML & ",""TD"""

    recipe = prefix & ",1,"">Earnings"",,,,1,,4)"
    smfMacAssertRecipe recipe, "May 02 AMC": checks = checks + 1
    smfMacAssertRecipe prefix & ")", "Earnings": checks = checks + 1
    smfMacAssertRecipe "smfGetTagContent(" & quotedHTML & ",""TD"",,,)", "Earnings"
    checks = checks + 1
    smfMacAssertRecipe prefix & ",-1,"">Price"")", "Price": checks = checks + 1
    smfMacAssertRecipe prefix & ",+1,"">Earnings"",,,,0,,0,4,"" "")", "<b>M"
    checks = checks + 1
    smfMacAssertRecipe prefix & ",1E0,"">Earnings"",,,,1)", "May 02 AMC"
    checks = checks + 1
    html = "<table><tr><td>Earnings, ""next)""</td><td>Quoted value</td></tr></table>"
    recipe = " = SmFgEtTaGcOnTeNt (" & smfMacFixtureQuote(html) & _
        ",""TD"",1," & smfMacFixtureQuote(">Earnings, ""next)""") & ") "
    smfMacAssertRecipe recipe, "Quoted value": checks = checks + 1
    recipe = prefix & ",1,"">Missing"",,,,1," & smfMacFixtureQuote("Missing, ""quoted)""") & ")"
    smfMacAssertRecipe recipe, "Missing, ""quoted)""": checks = checks + 1
    recipe = prefix & ",1,"">Missing"",,,,1,-77)"
    If Not smfMacTryTagContentRecipe(recipe, actual) Then Err.Raise 5, , "Numeric error literal was rejected"
    If Not IsNumeric(actual) Then Err.Raise 5, , "Numeric pError changed to text"
    If CDbl(actual) <> -77 Then Err.Raise 5, , "Numeric pError changed value"
    checks = checks + 1
    smfMacAssertRecipe prefix & ",1,"">Missing"")", "Error": checks = checks + 1
    smfMacAssertRecipe prefix & ",40000,"">Earnings"",,,,1,""tag overflow"")", "tag overflow"
    checks = checks + 1
    smfMacAssertRecipe prefix & ",1,"">Earnings"",,,,1,""length overflow"",0,40000)", "length overflow"
    checks = checks + 1

    rejected = Array( _
        prefix & ",1,"">Earnings"")*1000", _
        "=smfGetTagContent(LEFT(" & quotedHTML & ",10),""TD"")", _
        prefix & ",1+0,"">Earnings"")", _
        prefix & ",1,"">Earnings""", _
        prefix & ",1,"">Earnings)", _
        prefix & ",1,"">Earnings"",,,,,,,, ,0)", _
        "=smfGetTagContent(,""TD"")", _
        "=OtherFunction(" & quotedHTML & ",""TD"")")
    For Each item In rejected
        actual = "unchanged"
        If smfMacTryTagContentRecipe(CStr(item), actual) Then
            Err.Raise 5, , "Unsupported recipe was dispatched: " & CStr(item)
        End If
        If CStr(actual) <> "unchanged" Then Err.Raise 5, , "Rejected recipe changed its output argument"
        checks = checks + 1
    Next item
    If SMFMacLastURL() <> beforeURL Or SMFMacLastError() <> beforeError Or _
       SMFMacHTTPStatus() <> beforeStatus Then Err.Raise 5, , "Recipe fixtures changed HTTP state"
    SMFMacRecipeSelfTest = checks
End Function

Private Sub smfMacAssertRecipe(ByVal pRecipe As String, ByVal pExpected As String)
    Dim actual As Variant
    If Not smfMacTryTagContentRecipe(pRecipe, actual) Then
        Err.Raise 5, , "Literal recipe was rejected: " & pRecipe
    End If
    If CStr(actual) <> pExpected Then
        Err.Raise 5, , "Recipe expected [" & pExpected & "]; got [" & CStr(actual) & "]"
    End If
End Sub

Private Function smfMacFixtureQuote(ByVal pText As String) As String
    smfMacFixtureQuote = Chr$(34) & Replace(pText, Chr$(34), Chr$(34) & Chr$(34)) & Chr$(34)
End Function

' Separate from the existing 20 tag-recipe checks and 25 integration checks.
' Seed the normal page cache so every accepted call uses the real extractor.
Public Function SMFMacTableRecipeSelfTest() As Long
    Dim html As String, fixtureURL As String, quotedURL As String, prefix As String
    Dim actual As Variant, rejected As Variant, item As Variant, checks As Long, i As Long
    Dim oldKeys(1 To kPages) As String, oldBodies(1 To kPages) As String
    Dim oldTexts(1 To 3) As String, oldCacheMode As String, oldLog As String, oldError As Variant
    Dim beforeURL As String, beforeError As String, beforeStatus As Long, savedState As Boolean
    Dim errorNumber As Long, errorText As String
    On Error GoTo TestFailed
    beforeURL = SMFMacLastURL()
    beforeError = SMFMacLastError()
    beforeStatus = SMFMacHTTPStatus()
    oldCacheMode = sWebCache
    oldLog = sLog
    oldError = vError
    For i = 1 To kPages
        oldKeys(i) = aData(i, 1)
        oldBodies(i) = aData(i, 2)
    Next i
    For i = 1 To 3: oldTexts(i) = sData(i): Next i
    savedState = True
    Erase aData
    sWebCache = "Y"
    sLog = "N"
    fixtureURL = "fixture://table?arg=(one,two)"
    quotedURL = smfMacFixtureQuote(fixtureURL)
    prefix = "=RCHGetTableCell(" & quotedURL
    html = "<html><body><table><tr><td>Label</td><td>42</td><td></td><td>84</td></tr>" & _
        "<tr><td>Row Two</td><td>123.456789</td><td>Marker, ""next)""</td><td>Quoted value</td></tr>" & _
        "<tr><td>Row Three</td><td>126</td><td>Last</td><td>789</td></tr></table></body></html>"
    aData(1, 1) = "0:" & fixtureURL
    aData(1, 2) = html
    aData(2, 1) = "4:" & fixtureURL
    aData(2, 2) = Replace(html, ">42<", ">64<")

    smfMacAssertTableRecipe prefix & ",1)", "Label": checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1,"">Label"")", 42#: checks = checks + 1
    smfMacAssertTableRecipe " rChGeTtAbLeCeLl (" & quotedURL & ",1,"">Label"") ", 42#
    checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1," & smfMacFixtureQuote(">Marker, ""next)""") & ")", "Quoted value"
    checks = checks + 1
    smfMacAssertTableRecipe prefix & ",+1,"">Label"")", 42#: checks = checks + 1
    smfMacAssertTableRecipe prefix & ",-1,"">Label"")", 84#: checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1E0,"">Label"")", 42#: checks = checks + 1
    smfMacAssertTableRecipe prefix & ",.1E+1,"">Label"")", 42#: checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1,"">Label"",,,,,,,,)", 42#: checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1,"">Label"","">Row Two"")", 123.456789#
    checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1,"">Label"","">Row Two""," & _
        smfMacFixtureQuote(">Marker, ""next)""") & ")", "Quoted value"
    checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1,"">Label"",,,""missing|>Row Three"")", 126#
    checks = checks + 1
    smfMacAssertTableRecipe prefix & ",2,"">Label"",,,,1)", 123.456789#: checks = checks + 1
    smfMacAssertTableRecipe prefix & ",2,"">Row Three"",,,,-2)", 123.456789#: checks = checks + 1
    smfMacAssertTableRecipe prefix & ",2,"">Label"",,,,1,""</TR"",,""row limit"")", "row limit"
    checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1,"">42"",,,,,,1)", 84#: checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1,"">Label"",,,,,,,,4)", 64#: checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1,"">Missing"",,,,,,," & _
        smfMacFixtureQuote("Missing, ""quoted)""") & ")", "Missing, ""quoted)"""
    checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1,"">Missing"",,,,,,,-77)", -77#: checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1,"">Missing"",,,,,,,False)", False: checks = checks + 1
    smfMacAssertTableRecipe prefix & ",40000,"">Label"",,,,,,,""cell overflow"")", "cell overflow"
    checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1,"">Label"",,,,40000,,,""row overflow"")", "row overflow"
    checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1,"">Label"",,,,,,40000,""look overflow"")", "look overflow"
    checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1,"">Label"",,,,,,,""type overflow"",40000)", "type overflow"
    checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1E999,"">Label"",,,,,,,""number overflow"")", "number overflow"
    checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1,1E999,,,,,,,""find overflow"")", "find overflow"
    checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1,"">Missing"",,,,,,,+1E1)", 10#: checks = checks + 1
    smfMacAssertTableRecipe prefix & ",1,"">Missing"",,,,,,,True)", True: checks = checks + 1
    smfMacAssertTableRecipe prefix & ",""40000"","">Label"",,,,,,,""string overflow"")", "string overflow"
    checks = checks + 1
    smfMacAssertTableRecipe prefix & ",2,42)", 84#: checks = checks + 1

    rejected = Array( _
        prefix & ",1)*1000", _
        "=RCHGetTableCell(LEFT(" & quotedURL & ",10),1)", _
        prefix & ",1+0)", _
        prefix & ",1", _
        prefix & ",1,"">Label)", _
        prefix & ",1,,,,,,,,,,)", _
        "=RCHGetTableCell(,1)", _
        prefix & ",)", _
        "=OtherFunction(" & quotedURL & ",1)", _
        prefix & ",A1)")
    For Each item In rejected
        actual = "unchanged"
        If smfMacTryTableCellRecipe(CStr(item), actual) Then
            Err.Raise 5, , "Unsupported table recipe was dispatched: " & CStr(item)
        End If
        If CStr(actual) <> "unchanged" Then Err.Raise 5, , "Rejected table recipe changed its output argument"
        checks = checks + 1
    Next item
    If checks <> 40 Then Err.Raise 5, , "Unexpected table recipe fixture count"
    If SMFMacLastURL() <> beforeURL Or SMFMacLastError() <> beforeError Or _
       SMFMacHTTPStatus() <> beforeStatus Then Err.Raise 5, , "Table recipe fixtures changed HTTP state"
    SMFMacTableRecipeSelfTest = checks
    GoTo RestoreState
TestFailed:
    errorNumber = Err.Number
    errorText = Err.Description
    Resume RestoreState
RestoreState:
    On Error Resume Next
    If savedState Then
        For i = 1 To kPages
            aData(i, 1) = oldKeys(i)
            aData(i, 2) = oldBodies(i)
        Next i
        For i = 1 To 3: sData(i) = oldTexts(i): Next i
        sWebCache = oldCacheMode
        sLog = oldLog
        vError = oldError
    End If
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, "SMFMacTableRecipeSelfTest", errorText
End Function

Private Sub smfMacAssertTableRecipe(ByVal pRecipe As String, ByVal pExpected As Variant)
    Dim actual As Variant
    If Not smfMacTryTableCellRecipe(pRecipe, actual) Then
        Err.Raise 5, , "Literal table recipe was rejected: " & pRecipe
    End If
    If VarType(actual) <> VarType(pExpected) Then
        Err.Raise 5, , "Table recipe returned " & TypeName(actual) & "; expected " & TypeName(pExpected) & ": " & pRecipe
    End If
    If VarType(pExpected) = vbDouble Then
        If Abs(CDbl(actual) - CDbl(pExpected)) > 0.000000001 Then
            Err.Raise 5, , "Table recipe numeric value changed: " & pRecipe
        End If
    ElseIf CStr(actual) <> CStr(pExpected) Then
        Err.Raise 5, , "Table recipe expected [" & CStr(pExpected) & "]; got [" & CStr(actual) & "]: " & pRecipe
    End If
End Sub

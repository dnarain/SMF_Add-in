Attribute VB_Name = "modGetEconData"

#If Mac Then
Private mFREDKey As String
Private mFREDLiveURLs As Object
#End If

Function smfGetEconData(pID As String, _
                        pDate As Variant, _
               Optional pError As Variant = "Error")
    
    '-----------------------------------------------------------------------------------------------------------*
    ' Function to grab economic data from the St. Louis Federal Reserve web site
    '-----------------------------------------------------------------------------------------------------------*
    ' 2007.11.19 -- Created
    ' 2010.03.13 -- Add extraction of title elements
    ' 2011.04.28 -- Change cDec() to smfConvertData()
    ' 2016.06.20 -- Change to new URL of "https://fred.stlouisfed.org/data/"
    ' 2024.07.01 -- Web page change. Extract data from html instead of ".txt" format.
    '-----------------------------------------------------------------------------------------------------------*
    ' Example of usage:
    '
    '        =smfGetEconData("CURRNS", DATE(2007,3,24))
    '-----------------------------------------------------------------------------------------------------------*
    
#If Mac Then
    ' macOS uses the official API because the public data/download routes time out.
    ' Keep the key in smf-FRED-API-Key.txt beside this workbook, never in a formula.
    On Error GoTo MacError
    smfGetEconData = smfMacFREDValue(pID, pDate, pError)
    Exit Function
MacError:
    smfGetEconData = pError
#Else
    Dim sURL As String, sDataE As String
    On Error GoTo ErrorExit

    '--------------------------------> See if web page has already been retrieved
'   sURL = "https://research.stlouisfed.org/fred2/data/" & UCase(pID) & ".txt"
'   sURL = "https://fred.stlouisfed.org/data/" & UCase(pID) & ".txt"       ' New URL as of 2016-06-20
    sURL = "https://fred.stlouisfed.org/data/" & UCase(pID)                ' New URL as of 2024-07-01
    
    If IsDate(pDate) Then
        pDate = Format(pDate, "yyyy-mm-dd")
        End If
        
    smfGetEconData = RCHGetTableCell(sURL, 1, ">" & pDate & "<")
    Exit Function

ErrorExit:
    smfGetEconData = pError
    
#End If
    End Function


#If Mac Then
Private Function smfMacFREDValue(ByVal pID As String, ByRef pDate As Variant, ByVal pError As Variant) As Variant
    Dim id As String, label As String, iso As String, route As String, value As String
    Dim doc As Object, rows As Object, row As Object, source As Variant, answer As String
    Dim number As Double
    On Error GoTo Failed
    smfMacFREDValue = pError
    id = UCase$(Trim$(pID))
    If Len(id) = 0 Or Len(id) > 255 Or id Like "*[!A-Z0-9_.-]*" Then Exit Function
    If smfMacFREDDate(pDate, iso) Then
        ' Exact-date lookup: do not substitute an earlier observation on holidays.
        pDate = iso
        route = "series/observations?series_id=" & id & "&file_type=json&observation_start=" & iso & _
                "&observation_end=" & iso & "&sort_order=desc&limit=1"
        Set doc = smfMacFREDRequest(route)
        Set rows = doc("observations")
        If rows.Count <> 1 Then Exit Function
        Set row = rows(1)
        If CStr(row("date")) <> iso Then Exit Function
        If VarType(row("value")) <> vbString Then Exit Function
        value = Trim$(CStr(row("value")))
        If value = "." Then
            smfMacFREDValue = "."
        ElseIf smfMacFREDNumber(value, number) Then
            smfMacFREDValue = number
        End If
        Exit Function
    End If
    If IsObject(pDate) Or IsArray(pDate) Or IsError(pDate) Or IsNull(pDate) Then Exit Function
    label = LCase$(Trim$(CStr(pDate)))
    If Right$(label, 1) = ":" Then label = Trim$(Left$(label, Len(label) - 1))
    Select Case label
        Case "title", "series id", "seasonal adjustment", "frequency", "units", "date range", "last updated", "notes", "date"
            Set doc = smfMacFREDRequest("series?series_id=" & id & "&file_type=json")
            Set rows = doc("seriess")
            If rows.Count <> 1 Then Exit Function
            Set row = rows(1)
            If UCase$(smfMacFREDText(row, "id")) <> id Then Exit Function
            Select Case label
                Case "title": answer = smfMacFREDText(row, "title")
                Case "series id": answer = smfMacFREDText(row, "id")
                Case "seasonal adjustment": answer = smfMacFREDText(row, "seasonal_adjustment")
                Case "frequency": answer = smfMacFREDText(row, "frequency")
                Case "units": answer = smfMacFREDText(row, "units")
                Case "date range"
                    value = smfMacFREDText(row, "observation_start")
                    If Not smfMacFREDDate(value, iso) Then Exit Function
                    If value <> iso Then Exit Function
                    answer = iso & " to "
                    value = smfMacFREDText(row, "observation_end")
                    If Not smfMacFREDDate(value, iso) Then Exit Function
                    If value <> iso Then Exit Function
                    answer = answer & iso
                Case "last updated": answer = smfMacFREDText(row, "last_updated")
                Case "notes": answer = smfMacFREDText(row, "notes", True)
                Case "date": answer = "VALUE"
            End Select
        Case "release", "source"
            Set doc = smfMacFREDRequest("series/release?series_id=" & id & "&file_type=json")
            Set rows = doc("releases")
            If rows.Count <> 1 Then Exit Function
            Set row = rows(1)
            If label = "release" Then
                answer = smfMacFREDText(row, "name")
            Else
                If Not row.Exists("id") Then Exit Function
                If Not IsNumeric(row("id")) Then Exit Function
                If CDbl(row("id")) <= 0 Or CDbl(row("id")) <> Fix(CDbl(row("id"))) Then Exit Function
                Set doc = smfMacFREDRequest("release/sources?release_id=" & Format$(CDbl(row("id")), "0") & "&file_type=json")
                Set rows = doc("sources")
                If rows.Count = 0 Then Exit Function
                For Each source In rows
                    If Len(answer) > 0 Then answer = answer & "; "
                    Set row = source
                    answer = answer & smfMacFREDText(row, "name")
                Next source
            End If
        Case Else
            Exit Function
    End Select
    smfMacFREDValue = answer
Failed:
End Function

Private Function smfMacFREDText(ByVal pObject As Object, ByVal pField As String, _
                                Optional ByVal pAllowEmpty As Boolean = False) As String
    If Not pObject.Exists(pField) Then Err.Raise 5, "smfGetEconData", "Missing FRED field"
    If VarType(pObject(pField)) <> vbString Then Err.Raise 5, "smfGetEconData", "Invalid FRED field type"
    If Not pAllowEmpty And Len(pObject(pField)) = 0 Then Err.Raise 5, "smfGetEconData", "Empty FRED field"
    ' The shared decoder expands JSON \n to CRLF; remove that added CR once.
    smfMacFREDText = Replace(CStr(pObject(pField)), vbCrLf, vbLf)
End Function

Private Function smfMacFREDDate(ByVal pValue As Variant, ByRef pISO As String) As Boolean
    Dim text As String, candidate As Date, yearNumber As Long, monthNumber As Long, dayNumber As Long
    On Error GoTo Failed
    If IsObject(pValue) Or IsArray(pValue) Or IsError(pValue) Or IsNull(pValue) Or IsEmpty(pValue) Then Exit Function
    If VarType(pValue) = vbString Then
        text = Trim$(CStr(pValue))
        If Len(text) = 10 And Mid$(text, 5, 1) = "-" And Mid$(text, 8, 1) = "-" Then
            If Left$(text, 4) Like "*[!0-9]*" Or Mid$(text, 6, 2) Like "*[!0-9]*" Or Right$(text, 2) Like "*[!0-9]*" Then Exit Function
            yearNumber = CLng(Left$(text, 4)): monthNumber = CLng(Mid$(text, 6, 2)): dayNumber = CLng(Right$(text, 2))
            candidate = DateSerial(yearNumber, monthNumber, dayNumber)
            If Year(candidate) <> yearNumber Or Month(candidate) <> monthNumber Or Day(candidate) <> dayNumber Then Exit Function
        Else
            If Not IsDate(text) Then Exit Function
            candidate = CDate(text)
        End If
    ElseIf VarType(pValue) = vbDate Or IsNumeric(pValue) Then
        If VarType(pValue) = vbBoolean Then Exit Function
        candidate = CDate(pValue)
    Else
        Exit Function
    End If
    pISO = Format$(candidate, "yyyy-mm-dd")
    smfMacFREDDate = True
Failed:
End Function

Private Function smfMacFREDNumber(ByVal pText As String, ByRef pNumber As Double) As Boolean
    Dim pos As Long, digits As Long, length As Long, character As String
    On Error GoTo Failed
    length = Len(pText): pos = 1
    If length = 0 Then Exit Function
    character = Mid$(pText, pos, 1)
    If character = "+" Or character = "-" Then pos = pos + 1
    Do While pos <= length
        character = Mid$(pText, pos, 1)
        If character < "0" Or character > "9" Then Exit Do
        digits = digits + 1: pos = pos + 1
    Loop
    If Mid$(pText, pos, 1) = "." Then
        pos = pos + 1
        Do While pos <= length
            character = Mid$(pText, pos, 1)
            If character < "0" Or character > "9" Then Exit Do
            digits = digits + 1: pos = pos + 1
        Loop
    End If
    If digits = 0 Then Exit Function
    character = LCase$(Mid$(pText, pos, 1))
    If character = "e" Then
        pos = pos + 1: digits = 0
        character = Mid$(pText, pos, 1)
        If character = "+" Or character = "-" Then pos = pos + 1
        Do While pos <= length
            character = Mid$(pText, pos, 1)
            If character < "0" Or character > "9" Then Exit Do
            digits = digits + 1: pos = pos + 1
        Loop
        If digits = 0 Then Exit Function
    End If
    If pos <> length + 1 Then Exit Function
    ' Val reads the API's dot decimal independently of the Excel locale.
    pNumber = CDbl(Val(pText))
    smfMacFREDNumber = True
Failed:
End Function

Private Function smfMacFREDRequest(ByVal pRoute As String) As Object
    Const apiBase As String = "https://api.stlouisfed.org/fred/"
    Dim url As String, body As String, key As String, cacheKey As String
    Dim slot As Long, i As Long, oldURL As Variant, doc As Object, owned As Boolean
    On Error GoTo Failed
    url = apiBase & pRoute
    cacheKey = "0:" & url
    If mFREDLiveURLs Is Nothing Then Set mFREDLiveURLs = New Dictionary
    owned = mFREDLiveURLs.Exists(url)
    For i = 1 To kPages
        If aData(i, 1) = cacheKey Then slot = i: Exit For
    Next i
    ' Existing keyless cache entries also support offline callers without a key.
    If slot > 0 And sWebCache <> "N" And Not owned Then
        body = aData(slot, 2)
    Else
        key = smfMacFREDReadKey()
        If Len(key) = 0 Then Exit Function
        If key <> mFREDKey Then
            ' Never retain data fetched with another local key after configuration changes.
            For i = 1 To kPages
                If Left$(aData(i, 1), Len("0:" & apiBase)) = "0:" & apiBase Then
                    oldURL = Mid$(aData(i, 1), 3)
                    If mFREDLiveURLs.Exists(CStr(oldURL)) Then
                        aData(i, 1) = "": aData(i, 2) = ""
                    End If
                End If
            Next i
            mFREDLiveURLs.RemoveAll
            mFREDKey = key
            slot = 0
        End If
        If slot > 0 And sWebCache <> "N" Then
            body = aData(slot, 2)
        Else
            ' smfMacFetch redacts api_key from diagnostics. Do not pass this URL to the generic page logger.
            body = smfMacFetch(url & "&api_key=" & key)
            If Left$(body, 9) = "Error -- " Then Exit Function
        End If
    End If
    Set doc = JsonConverter.ParseJson(body)
    If doc.Exists("error_code") Or doc.Exists("error_message") Then Exit Function
    If Len(key) > 0 Then
        If slot = 0 Then
            For i = 1 To kPages
                If Len(aData(i, 1)) = 0 Then slot = i: Exit For
            Next i
        End If
        If slot > 0 Then
            aData(slot, 1) = cacheKey
            aData(slot, 2) = body
            mFREDLiveURLs(url) = True
        End If
    End If
    Set smfMacFREDRequest = doc
Failed:
End Function

Private Function smfMacFREDReadKey() As String
    Dim fileNumber As Integer, key As String, opened As Boolean
    On Error GoTo Failed
    fileNumber = FreeFile
    Open ThisWorkbook.Path & "/smf-FRED-API-Key.txt" For Input As #fileNumber
    opened = True
    Line Input #fileNumber, key
    key = Trim$(key)
    If Len(key) = 32 And Not key Like "*[!a-z0-9]*" Then smfMacFREDReadKey = key
Failed:
    If opened Then Close #fileNumber
End Function
#End If

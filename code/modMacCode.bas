Attribute VB_Name = "modMacCode"
Option Explicit

' Native Excel for Mac transport. Install mac/SMFMac.applescript in
' ~/Library/Application Scripts/com.microsoft.Excel/ before use.
' Request: METHOD + LF + URL. Response: SMFHTTP/1 + LF + status + LF + body.
Public cookieArr As Variant
Public iCookieInit As Integer
Private mHTTPStatus As Long
Private mHTTPError As String
Private mHTTPURL As String
Private mFileError As String
#If Mac Then
Private mMacFetchCalls As Double
Private mMacFetchRequests As Double
#End If

Public Function RCHGetURLData1Mac(ByVal pURL As String, _
                                Optional ByVal pUseIE As Integer = 0) As String
    Dim method As String
    method = "GET"
    If pUseIE = 3 Then method = "POST"
    If pUseIE = 4 Then method = "YAHOO"
    RCHGetURLData1Mac = smfMacFetch(pURL, method)
End Function

Public Function smfMacFetch(ByVal pURL As String, Optional ByVal pMethod As String = "GET") As String
    Dim response As String, errorNumber As Long, errorText As String
#If Mac Then
    mMacFetchCalls = mMacFetchCalls + 1#
#End If
    mHTTPStatus = 0
    mHTTPError = vbNullString
#If Mac Then
    mHTTPURL = smfMacRedactAPIKey(pURL)
#Else
    mHTTPURL = pURL
#End If
    On Error GoTo RequestError

    If InStr(pURL, vbCr) > 0 Or InStr(pURL, vbLf) > 0 Then
        Err.Raise vbObjectError + 2100, "smfMacFetch", "A URL cannot contain a line break"
    End If
    If LCase$(Left$(pURL, 7)) <> "http://" And LCase$(Left$(pURL, 8)) <> "https://" Then
        Err.Raise vbObjectError + 2101, "smfMacFetch", "Only HTTP and HTTPS URLs are supported"
    End If
    pMethod = UCase$(pMethod)
    If pMethod <> "GET" And pMethod <> "POST" And pMethod <> "YAHOO" Then
        Err.Raise vbObjectError + 2102, "smfMacFetch", "Unsupported request method"
    End If
#If Mac Then
    Dim unsupportedID As String, unsupportedReason As String
    Call smfUnsupportedURL(pURL, unsupportedID, unsupportedReason)
    If Len(unsupportedID) > 0 Then
        mHTTPError = "Not Supported -- " & unsupportedReason
        smfMacFetch = "Error -- " & mHTTPError
        Exit Function
    End If
    mMacFetchRequests = mMacFetchRequests + 1#
    response = AppleScriptTask("SMFMac.scpt", "fetchURL", pMethod & vbLf & pURL)
    smfMacFetch = smfMacReadResponse(response)
#Else
    Err.Raise vbObjectError + 2103, "smfMacFetch", "This function requires Excel for Mac"
#End If
    Exit Function

RequestError:
    errorNumber = Err.Number
    errorText = Err.Description
#If Mac Then
    errorText = smfMacRedactAPIKey(errorText)
#End If
    mHTTPError = "Mac request failed (" & CStr(errorNumber) & "): " & errorText
    smfMacFetch = "Error -- " & mHTTPError
End Function

#If Mac Then
' Keep API-key query values out of saved request/error diagnostics.
Private Function smfMacRedactAPIKey(ByVal pText As String) As String
    Dim view As String, starts() As Long, pos As Long, characterWidth As Long
    Dim count As Long, high As Long, low As Long, credentialName As Variant, nameEnd As Long
    Dim operatorAt As Long, valueAt As Long, endAt As Long, rawAt As Long, back As Long
    Dim encoded As Boolean, rawDelimiter As Boolean, quote As String, ch As String, result As String
    On Error GoTo Failed
    If Len(pText) = 0 Then Exit Function
    ReDim starts(1 To Len(pText) + 1)
    view = Space$(Len(pText))
    pos = 1
    Do While pos <= Len(pText)
        ch = Mid$(pText, pos, 1)
        characterWidth = 1
        If ch = "%" And pos + 2 <= Len(pText) Then
            high = InStr(1, "0123456789abcdef", LCase$(Mid$(pText, pos + 1, 1)), vbBinaryCompare) - 1
            low = InStr(1, "0123456789abcdef", LCase$(Mid$(pText, pos + 2, 1)), vbBinaryCompare) - 1
            If high >= 0 And low >= 0 And high * 16 + low < 128 Then
                ch = Chr$(high * 16 + low)
                characterWidth = 3
            End If
        End If
        count = count + 1
        starts(count) = pos
        Mid$(view, count, 1) = ch
        pos = pos + characterWidth
    Loop
    starts(count + 1) = Len(pText) + 1
    view = Left$(view, count)
    pos = 1
    rawAt = 1
    Do While pos <= count
        If pos > 1 Then
            ch = Mid$(view, pos - 1, 1)
            If InStr(1, "?&;#{}[, (" & vbTab & vbCr & vbLf & Chr$(34) & "'", ch, vbBinaryCompare) = 0 Then GoTo NextCharacter
        End If
        For Each credentialName In Array("api_key", "apikey", "api-key")
            If StrComp(Mid$(view, pos, Len(CStr(credentialName))), CStr(credentialName), vbTextCompare) = 0 Then Exit For
        Next credentialName
        If IsEmpty(credentialName) Then GoTo NextCharacter
        ' For Each retains its final item after a failed search; confirm the match.
        If StrComp(Mid$(view, pos, Len(CStr(credentialName))), CStr(credentialName), vbTextCompare) <> 0 Then GoTo NextCharacter
        nameEnd = pos + Len(CStr(credentialName))
        If nameEnd <= count Then
            ch = LCase$(Mid$(view, nameEnd, 1))
            If InStr(1, "abcdefghijklmnopqrstuvwxyz0123456789_-", ch, vbBinaryCompare) > 0 Then GoTo NextCharacter
        End If
        operatorAt = nameEnd
        ch = Mid$(view, operatorAt, 1)
        If ch = Chr$(34) Or ch = "'" Then operatorAt = operatorAt + 1
        Do While operatorAt <= count
            ch = Mid$(view, operatorAt, 1)
            If InStr(1, " " & vbTab & vbCr & vbLf, ch, vbBinaryCompare) = 0 Then Exit Do
            operatorAt = operatorAt + 1
        Loop
        If operatorAt > count Then GoTo NextCharacter
        ch = Mid$(view, operatorAt, 1)
        If ch <> ":" And ch <> "=" Then GoTo NextCharacter
        encoded = starts(operatorAt + 1) - starts(operatorAt) > 1
        If pos > 1 Then
            If starts(pos) - starts(pos - 1) > 1 Then encoded = True
        End If
        valueAt = operatorAt + 1
        Do While valueAt <= count
            ch = Mid$(view, valueAt, 1)
            If InStr(1, " " & vbTab & vbCr & vbLf, ch, vbBinaryCompare) = 0 Then Exit Do
            valueAt = valueAt + 1
        Loop
        quote = ""
        ch = Mid$(view, valueAt, 1)
        If ch = Chr$(34) Or ch = "'" Then
            quote = ch
            If starts(valueAt + 1) - starts(valueAt) > 1 Then encoded = True
            valueAt = valueAt + 1
        End If
        endAt = valueAt
        Do While endAt <= count
            ch = Mid$(view, endAt, 1)
            rawDelimiter = starts(endAt + 1) - starts(endAt) = 1 Or encoded
            If Len(quote) > 0 Then
                If ch = quote And rawDelimiter Then
                    back = endAt - 1
                    Do While back >= valueAt
                        If Mid$(view, back, 1) <> "\" Then Exit Do
                        back = back - 1
                    Loop
                    If (endAt - back - 1) Mod 2 = 0 Then Exit Do
                End If
            ElseIf rawDelimiter Then
                If InStr(1, "&;# " & vbTab & vbCr & vbLf & Chr$(34) & "'<>},])", ch, vbBinaryCompare) > 0 Then Exit Do
            End If
            endAt = endAt + 1
        Loop
        If endAt > valueAt Then
            result = result & Mid$(pText, rawAt, starts(valueAt) - rawAt) & "[redacted]"
            rawAt = starts(endAt)
        End If
        pos = endAt
        GoTo ContinueScan
NextCharacter:
        pos = pos + 1
ContinueScan:
    Loop
    smfMacRedactAPIKey = result & Mid$(pText, rawAt)
    Exit Function
Failed:
    smfMacRedactAPIKey = "[redacted diagnostic]"
End Function
#End If

' Public for deterministic tests. Preserve line endings in the response body.
Public Function smfMacReadResponse(ByVal pResponse As String) As String
    Dim firstLF As Long, secondLF As Long, statusText As String, body As String, i As Long
    mHTTPStatus = 0
    mHTTPError = vbNullString
    firstLF = InStr(1, pResponse, vbLf, vbBinaryCompare)
    If firstLF = 0 Then GoTo InvalidResponse
    If Left$(pResponse, firstLF - 1) <> "SMFHTTP/1" Then GoTo InvalidResponse
    secondLF = InStr(firstLF + 1, pResponse, vbLf, vbBinaryCompare)
    If secondLF = 0 Then GoTo InvalidResponse
    statusText = Mid$(pResponse, firstLF + 1, secondLF - firstLF - 1)
    If Len(statusText) = 0 Or Len(statusText) > 3 Then GoTo InvalidResponse
    For i = 1 To Len(statusText)
        If Mid$(statusText, i, 1) < "0" Or Mid$(statusText, i, 1) > "9" Then GoTo InvalidResponse
    Next i
    mHTTPStatus = CLng(statusText)
    body = Mid$(pResponse, secondLF + 1)
    If mHTTPStatus >= 200 And mHTTPStatus <= 299 Then
        smfMacReadResponse = body
    Else
        If mHTTPStatus = 0 Then
#If Mac Then
            mHTTPError = smfMacRedactAPIKey(Left$(body, 500))
#Else
            mHTTPError = Left$(body, 500)
#End If
        Else
            mHTTPError = "HTTP " & CStr(mHTTPStatus)
        End If
        smfMacReadResponse = "Error -- " & mHTTPError
    End If
    Exit Function
InvalidResponse:
    mHTTPError = "Invalid response from SMFMac.applescript"
    smfMacReadResponse = "Error -- " & mHTTPError
End Function

Public Function SMFMacHTTPStatus() As Long
    SMFMacHTTPStatus = mHTTPStatus
End Function

Public Function SMFMacLastError() As String
    SMFMacLastError = mHTTPError
End Function

Public Function SMFMacLastURL() As String
    SMFMacLastURL = mHTTPURL
End Function

#If Mac Then
' pRequests counts fetchURL bridge attempts, including attempts that fail in the helper.
Public Sub SMFMacRequestCounts(ByRef pCalls As Double, ByRef pRequests As Double)
    pCalls = mMacFetchCalls
    pRequests = mMacFetchRequests
End Sub
#End If

' Handle a complete literal smfGetTagContent call without Excel's evaluator.
' False means unsupported syntax; pValue is unchanged in that case.
Public Function smfMacTryTagContentRecipe(ByVal pExpression As String, ByRef pValue As Variant) As Boolean
    Dim expression As String, tokens(1 To 12) As String, values(1 To 12) As Variant
    Dim count As Long, i As Long, openAt As Long, tagCount As Integer
    Dim convert As Integer, requestType As Integer, maxLength As Integer
    On Error GoTo RecipeFailed
    expression = Trim$(pExpression)
    If Left$(expression, 1) = "=" Then expression = Trim$(Mid$(expression, 2))
    openAt = InStr(1, expression, "(", vbBinaryCompare)
    If openAt = 0 Then Exit Function
    If LCase$(Trim$(Left$(expression, openAt - 1))) <> "smfgettagcontent" Then Exit Function
    If Not smfMacRecipeArguments(Mid$(expression, openAt + 1), tokens, count) Then Exit Function
    If count < 2 Or Len(tokens(1)) = 0 Or Len(tokens(2)) = 0 Then Exit Function

    values(3) = 1
    values(4) = "<"
    values(5) = " "
    values(6) = " "
    values(7) = " "
    values(8) = 0
    values(9) = "Error"
    values(10) = 0
    values(11) = 32767
    values(12) = " "
    For i = 1 To count
        If Len(tokens(i)) > 0 Then
            If Not smfMacRecipeLiteral(tokens(i), values(i)) Then Exit Function
        End If
    Next i
    ' Accepted literal calls must not fall back to Evaluate on type overflow.
    smfMacTryTagContentRecipe = True
    tagCount = CInt(values(3))
    convert = CInt(values(8))
    requestType = CInt(values(10))
    maxLength = CInt(values(11))
    pValue = smfGetTagContent(CStr(values(1)), CStr(values(2)), tagCount, _
        CStr(values(4)), CStr(values(5)), CStr(values(6)), CStr(values(7)), _
        convert, values(9), requestType, maxLength, CStr(values(12)))
    Exit Function
RecipeFailed:
    ' Do not retry an accepted call through Evaluate after a runtime error.
    If smfMacTryTagContentRecipe Then pValue = values(9)
End Function

' Handle the complete literal RCHGetTableCell call without Excel's evaluator.
' Validate syntax before numeric conversion so overflows remain handled calls.
Public Function smfMacTryTableCellRecipe(ByVal pExpression As String, ByRef pValue As Variant) As Boolean
    Dim expression As String, tokens(1 To 11) As String, values(1 To 11) As Variant
    Dim count As Long, i As Long, openAt As Long
    Dim cells As Integer, rows As Integer, look As Integer, requestType As Integer
    On Error GoTo RecipeFailed
    expression = Trim$(pExpression)
    If Left$(expression, 1) = "=" Then expression = Trim$(Mid$(expression, 2))
    openAt = InStr(1, expression, "(", vbBinaryCompare)
    If openAt = 0 Then Exit Function
    If LCase$(Trim$(Left$(expression, openAt - 1))) <> "rchgettablecell" Then Exit Function
    If Not smfMacRecipeArguments(Mid$(expression, openAt + 1), tokens, count) Then Exit Function
    If count < 2 Or Len(tokens(1)) = 0 Or Len(tokens(2)) = 0 Then Exit Function

    values(3) = "<BODY"
    values(4) = " "
    values(5) = " "
    values(6) = " "
    values(7) = 0
    values(8) = "</BODY"
    values(9) = 0
    values(10) = "Error"
    values(11) = 0
    For i = 1 To count
        If Len(tokens(i)) > 0 Then
            ' Defer Val: even a valid numeric literal can overflow Double.
            If Not smfMacRecipeIsNumber(tokens(i)) Then
                If Not smfMacRecipeLiteral(tokens(i), values(i)) Then Exit Function
            End If
        End If
    Next i
    smfMacTryTableCellRecipe = True
    ' Read pError first, including when another numeric argument overflows.
    If Len(tokens(10)) > 0 Then
        If smfMacRecipeIsNumber(tokens(10)) Then values(10) = Val(tokens(10))
    End If
    For i = 1 To count
        If Len(tokens(i)) > 0 Then
            If smfMacRecipeIsNumber(tokens(i)) Then values(i) = Val(tokens(i))
        End If
    Next i
    cells = CInt(values(2))
    rows = CInt(values(7))
    look = CInt(values(9))
    requestType = CInt(values(11))
    pValue = RCHGetTableCell(CStr(values(1)), cells, CStr(values(3)), _
        CStr(values(4)), CStr(values(5)), CStr(values(6)), rows, _
        CStr(values(8)), look, values(10), requestType)
    Exit Function
RecipeFailed:
    If smfMacTryTableCellRecipe Then pValue = values(10)
End Function

#If Mac Then
' Accept only complete calls from this small function list. Validate the whole
' expression without fetching data before executing any approved VBA call.
Public Function smfMacTryNestedRecipe(ByVal pExpression As String, ByRef pValue As Variant, _
                                    Optional ByVal pError As Variant = "Error") As Boolean
    Dim expression As String, unused As Variant, calls As Long
    On Error GoTo RecipeFailed
    expression = Trim$(pExpression)
    If Left$(expression, 1) = "=" Then expression = Trim$(Mid$(expression, 2))
    If Len(expression) = 0 Or Len(expression) > 8192 Then Exit Function
    If Not smfMacNestedRecipeNode(expression, False, unused, 1, calls) Then Exit Function
    smfMacTryNestedRecipe = True
    calls = 0
    If Not smfMacNestedRecipeNode(expression, True, pValue, 1, calls) Then Err.Raise 5
    Exit Function
RecipeFailed:
    ' An accepted call must never retry through Excel's evaluator after an error.
    If smfMacTryNestedRecipe Then pValue = pError
End Function

Private Function smfMacNestedRecipeNode(ByVal pExpression As String, ByVal pExecute As Boolean, _
                                      ByRef pValue As Variant, ByVal pDepth As Long, _
                                      ByRef pCalls As Long) As Boolean
    Dim tokens(1 To 12) As String, values(1 To 12) As Variant
    Dim name As String, count As Long, required As Long, maximum As Long
    Dim i As Long, openAt As Long, canNest As Boolean
    If pDepth > 8 Or pCalls >= 16 Then Exit Function
    pCalls = pCalls + 1
    openAt = InStr(1, pExpression, "(", vbBinaryCompare)
    If openAt = 0 Then Exit Function
    name = LCase$(Trim$(Left$(pExpression, openAt - 1)))
    Select Case name
        Case "smfgettagcontent": required = 2: maximum = 12
            values(9) = "Error"
        Case "smfgetyahoojsonfield": required = 3: maximum = 4: values(4) = 0
        Case "smfconvertdata": required = 1: maximum = 2: values(2) = 0: canNest = True
        Case "smfstrextr": required = 3: maximum = 4: values(4) = 0: canNest = True
        Case "datevalue": required = 1: maximum = 1: canNest = True
        Case Else: Exit Function
    End Select
    If Not smfMacNestedArguments(Mid$(pExpression, openAt + 1), tokens, count) Then Exit Function
    If count < required Or count > maximum Then Exit Function
    For i = 1 To required
        If Len(tokens(i)) = 0 Then Exit Function
    Next i
    For i = 1 To count
        If Len(tokens(i)) > 0 Then
            If smfMacRecipeIsNumber(tokens(i)) Then
                ' Syntax validation must not coerce an overflowing number.
            ElseIf smfMacRecipeLiteral(tokens(i), values(i)) Then
                ' Strings and Boolean literals have no external effects.
            ElseIf i = 1 And canNest Then
                If Not smfMacNestedRecipeNode(tokens(i), pExecute, values(i), pDepth + 1, pCalls) Then Exit Function
            Else
                Exit Function
            End If
        End If
    Next i
    smfMacNestedRecipeNode = True
    If Not pExecute Then Exit Function

    If name = "smfgettagcontent" Then
        ' The existing literal dispatcher preserves all twelve tag arguments.
        ' Its False result on an overflowing numeric literal is handled here.
        If smfMacRecipeIsNumber(tokens(9)) Then values(9) = Val(tokens(9))
        If Not smfMacTryTagContentRecipe(pExpression, pValue) Then pValue = values(9)
        Exit Function
    End If
    For i = 1 To count
        If Len(tokens(i)) > 0 Then
            If smfMacRecipeIsNumber(tokens(i)) Then values(i) = Val(tokens(i))
        End If
    Next i
    Select Case name
        Case "smfgetyahoojsonfield"
            pValue = smfGetYahooJSONField(CStr(values(1)), CStr(values(2)), CStr(values(3)), CDbl(values(4)))
        Case "smfconvertdata"
            pValue = smfConvertData(CStr(values(1)), CInt(values(2)))
        Case "smfstrextr"
            ' smfStrExtr has ByRef parameters; pass explicit typed expressions.
            pValue = smfStrExtr(CStr(values(1)), CStr(values(2)), CStr(values(3)), CInt(values(4)))
        Case "datevalue"
            pValue = DateValue(CStr(values(1)))
    End Select
End Function

Private Function smfMacNestedArguments(ByVal pText As String, ByRef pTokens() As String, _
                                     ByRef pCount As Long) As Boolean
    Dim i As Long, startAt As Long, depth As Long, quoted As Boolean, character As String
    pCount = 1
    startAt = 1
    i = 1
    Do While i <= Len(pText)
        character = Mid$(pText, i, 1)
        If character = Chr$(34) Then
            If quoted And i < Len(pText) Then
                If Mid$(pText, i + 1, 1) = Chr$(34) Then
                    i = i + 2
                    GoTo NextCharacter
                End If
            End If
            quoted = Not quoted
        ElseIf Not quoted Then
            Select Case character
                Case "("
                    depth = depth + 1
                    If depth > 8 Then Exit Function
                Case ")"
                    If depth > 0 Then
                        depth = depth - 1
                    Else
                        pTokens(pCount) = Trim$(Mid$(pText, startAt, i - startAt))
                        If Len(Trim$(Mid$(pText, i + 1))) > 0 Then Exit Function
                        smfMacNestedArguments = True
                        Exit Function
                    End If
                Case ","
                    If depth = 0 Then
                        pTokens(pCount) = Trim$(Mid$(pText, startAt, i - startAt))
                        pCount = pCount + 1
                        If pCount > UBound(pTokens) Then Exit Function
                        startAt = i + 1
                    End If
            End Select
        End If
        i = i + 1
NextCharacter:
    Loop
End Function
#End If

Private Function smfMacRecipeArguments(ByVal pText As String, ByRef pTokens() As String, _
                                       ByRef pCount As Long) As Boolean
    Dim i As Long, startAt As Long, quoted As Boolean, character As String
    pCount = 1
    startAt = 1
    i = 1
    Do While i <= Len(pText)
        character = Mid$(pText, i, 1)
        If character = Chr$(34) Then
            If quoted And i < Len(pText) Then
                If Mid$(pText, i + 1, 1) = Chr$(34) Then
                    i = i + 2
                    GoTo NextCharacter
                End If
            End If
            quoted = Not quoted
        ElseIf Not quoted Then
            Select Case character
                Case "(": Exit Function
                Case ",", ")"
                    pTokens(pCount) = Trim$(Mid$(pText, startAt, i - startAt))
                    If character = ")" Then
                        If Len(Trim$(Mid$(pText, i + 1))) > 0 Then Exit Function
                        smfMacRecipeArguments = True
                        Exit Function
                    End If
                    pCount = pCount + 1
                    If pCount > UBound(pTokens) Then Exit Function
                    startAt = i + 1
            End Select
        End If
        i = i + 1
NextCharacter:
    Loop
End Function

Private Function smfMacRecipeLiteral(ByVal pToken As String, ByRef pValue As Variant) As Boolean
    Dim i As Long, decoded As String, character As String
    If Left$(pToken, 1) = Chr$(34) Then
        i = 2
        Do While i <= Len(pToken)
            character = Mid$(pToken, i, 1)
            If character = Chr$(34) Then
                If i = Len(pToken) Then
                    pValue = decoded
                    smfMacRecipeLiteral = True
                    Exit Function
                End If
                If Mid$(pToken, i + 1, 1) <> Chr$(34) Then Exit Function
                i = i + 1
            End If
            decoded = decoded & character
            i = i + 1
        Loop
    ElseIf LCase$(pToken) = "true" Or LCase$(pToken) = "false" Then
        pValue = (LCase$(pToken) = "true")
        smfMacRecipeLiteral = True
    ElseIf smfMacRecipeIsNumber(pToken) Then
        ' Val reads the formula's decimal point independently of locale.
        pValue = Val(pToken)
        smfMacRecipeLiteral = True
    End If
End Function

Private Function smfMacRecipeIsNumber(ByVal pToken As String) As Boolean
    Dim i As Long, digits As Long, exponentDigits As Long, character As String
    i = 1
    If Left$(pToken, 1) = "+" Or Left$(pToken, 1) = "-" Then i = i + 1
    Do While i <= Len(pToken)
        character = Mid$(pToken, i, 1)
        If character < "0" Or character > "9" Then Exit Do
        digits = digits + 1
        i = i + 1
    Loop
    If Mid$(pToken, i, 1) = "." Then
        i = i + 1
        Do While i <= Len(pToken)
            character = Mid$(pToken, i, 1)
            If character < "0" Or character > "9" Then Exit Do
            digits = digits + 1
            i = i + 1
        Loop
    End If
    If digits = 0 Then Exit Function
    If UCase$(Mid$(pToken, i, 1)) = "E" Then
        i = i + 1
        If Mid$(pToken, i, 1) = "+" Or Mid$(pToken, i, 1) = "-" Then i = i + 1
        Do While i <= Len(pToken)
            character = Mid$(pToken, i, 1)
            If character < "0" Or character > "9" Then Exit Do
            exponentDigits = exponentDigits + 1
            i = i + 1
        Loop
        If exponentDigits = 0 Then Exit Function
    End If
    smfMacRecipeIsNumber = (i > Len(pToken))
End Function

Public Function smfMacReadElementText(ByVal pFolder As String, ByVal pSuffix As String) As String
    mFileError = vbNullString
    On Error GoTo ReadFailed
#If Mac Then
    smfMacReadElementText = AppleScriptTask("SMFMac.scpt", "readElementFile", pFolder & vbLf & pSuffix)
#Else
    Err.Raise 5, , "This function requires Excel for Mac"
#End If
    Exit Function
ReadFailed:
    mFileError = "Element file read failed: " & Err.Description
    Err.Raise Err.Number, "smfMacReadElementText", mFileError
End Function

Public Function SMFMacLastFileError() As String
    SMFMacLastFileError = mFileError
End Function

Public Function smfMacReadElementLocation(ByVal pFolder As String) As String
    mFileError = vbNullString
    On Error GoTo ReadFailed
#If Mac Then
    smfMacReadElementLocation = AppleScriptTask("SMFMac.scpt", "readElementLocation", pFolder)
#Else
    Err.Raise 5, , "This function requires Excel for Mac"
#End If
    Exit Function
ReadFailed:
    mFileError = "Element location read failed: " & Err.Description
    Err.Raise Err.Number, "smfMacReadElementLocation", mFileError
End Function

Public Function SMFMacTransportSelfTest() As Long
    Dim actual As String
    actual = smfMacReadResponse("SMFHTTP/1" & vbLf & "200" & vbLf & "A" & vbLf & "B")
    If actual <> "A" & vbLf & "B" Then Err.Raise 5, , "HTTP response body changed"
    If SMFMacHTTPStatus() <> 200 Then Err.Raise 5, , "HTTP status was not retained"
    actual = smfMacReadResponse("SMFHTTP/1" & vbLf & "403" & vbLf & "<html>Denied</html>")
    If actual <> "Error -- HTTP 403" Then Err.Raise 5, , "HTTP error was treated as data"
    actual = smfMacReadResponse("SMFHTTP/1" & vbLf & "0" & vbLf & "Timed out")
    If SMFMacLastError() <> "Timed out" Then Err.Raise 5, , "Transport error was lost"
    actual = smfMacReadResponse("not a response")
    If SMFMacHTTPStatus() <> 0 Or SMFMacLastError() = "" Then Err.Raise 5, , "Malformed response accepted"
    actual = smfMacReadResponse("SMFHTTP/1" & vbLf & "200" & vbLf)
    If actual <> "" Or SMFMacLastError() <> "" Then Err.Raise 5, , "Empty successful response rejected"
    SMFMacTransportSelfTest = 6
End Function



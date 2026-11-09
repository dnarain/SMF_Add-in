Attribute VB_Name = "modGetAdvFNElement"
Public sAdvFNPrefix As String                       ' URL prefix to use for AdvFN
Function smfGetADVFNElement(ByVal pTicker As String, _
                            ByVal pPeriod As String, _
                            ByVal pCells As Integer, _
                   Optional ByVal pFind1 As String = "", _
                   Optional ByVal pFind2 As String = "", _
                   Optional ByVal pError As Variant = "Error", _
                   Optional ByVal pType As Integer = 0) As Variant
                        
   '-----------------------------------------------------------------------------------------------------------*
   ' Function to return a financial statements data element from AdvFN
   '-----------------------------------------------------------------------------------------------------------*
   ' 2009.11.04 -- Created by Randy Harmelink (rharmelink@gmail.com)
   ' 2009.12.19 -- Add smfConvertYahooTicker() processing
   ' 2012.01.30 -- Change AdvFn URL from "http://www..." to "http://us...."
   ' 2012.02.02 -- Make AdvFN URL prefix a variable
   ' 2012.02.04 -- Default AdvFN URL prefix to "www" unless external file exists to set it to something else
   ' 2013.10.16 -- Update for AdvFN URL changes
   ' 2013.10.17 -- Fix for above update
   ' 2014.12.02 -- Fix error processing when calling external functions
   ' 2015.03.05 -- Change "/exchanges/" search string to "/stock-market/"
   ' 2015.03.21 -- Allow exchange prefix on ticker symbol
   ' 2017.05.05 -- Add smfSetAdvFNPrefix() sub
   ' 2018.01.24 -- Change AdvFN URL from "http://" to "https://"
   ' 2023-09-24 -- https://climbermel.github.io/SMF_Add-in/
   '               5196 - 13625 AdvFN-A or Q all return error.  smf-elements-7.txt
   '               https://ca.advfn.com/stock-market/NASDAQ/AAPL/financials has data,
   '               but I doubt it is worth trying to fix any of these elements...
   '-----------------------------------------------------------------------------------------------------------*
   ' Sample of use:
   '
   '    =smfGetADVFNElement("MMM","A",1,">Year End Date")
   '
   '-----------------------------------------------------------------------------------------------------------*
    
    On Error GoTo ErrorExit
    vError = pError
#If Mac Then
    ' The current annual page writes the latest fiscal date in its header.
    ' Older periods and other data items still use the original extractor below.
    Dim macYearEndLabel As String
    macYearEndLabel = LCase$(Trim$(pFind1))
    If Left$(macYearEndLabel, 1) = ">" Then macYearEndLabel = Trim$(Mid$(macYearEndLabel, 2))
    If UCase$(pPeriod) = "A" And pCells = 1 And _
       (macYearEndLabel = "year end date" Or macYearEndLabel = "year end") Then
        smfGetADVFNElement = smfMacADVFNLatestYearEnd(pTicker, pFind2, pError, pType)
        Exit Function
    End If
#End If
    
    '------------------> Create labels for annual vs quarterly processing
    Dim sTicker As String, sExchange As String
    sTicker = smfConvertYahooTicker(pTicker, "ADVFN")
    sExchange = smfStrExtr(sTicker, "~", ":")
    If sExchange <> "" Then sTicker = smfStrExtr(sTicker, ":", "~")
    Select Case UCase(pPeriod)
       Case "A": sLabel1 = "&mode=annual_reports": sLabel2 = "start_date"
       Case "Q": sLabel1 = "&mode=quarterly_reports": sLabel2 = "istart_date"
       Case Else: vError = "Improper period -- should be A or Q": GoTo ErrorExit
       End Select
    If sAdvFNPrefix = "" Then LoadAdvFNPrefix
    If sExchange = "" Then
       sURL = "https://" & sAdvFNPrefix & ".advfn.com/p.php?pid=financials&mode=quarterly_reports&symbol=" & sTicker
       sExchange = smfStrExtr(RCHGetWebData(sURL, "/stock-market/", 50), "/stock-market/", "/")
       End If
    sURL = "https://" & sAdvFNPrefix & ".advfn.com/stock-market/" & sExchange & "/" & sTicker & "/financials?" _
         & sLabel1 & "&btn=" & sLabel2
    
    '------------------> Determine # of available periods and paging points (5 periods per page)
    nPeriods = smfConvertData(smfStrExtr(Right(smfGetTagContent(sURL, "select", 1, "Select start date", pError:=vError), 35), "'", "'"))
    If pCells = 999 Then
       smfGetADVFNElement = nPeriods + 1
       Exit Function
       End If
    If pCells > nPeriods + 1 Then GoTo ErrorExit
    nRawID = nPeriods - 5 * (Int((pCells - 1) / 5) + 1)
    nColumn = (200 - pCells) Mod 5 + 1
    Select Case nRawID
       Case Is < 0
            If nPeriods < 5 Then
               nPageId = ""
            Else
               nPageId = "&" & sLabel2 & "=0"
               End If
            nColumn = nColumn + nRawID + 1
       Case Is = nPeriods - 5
            nPageId = ""
       Case Else
            nPageId = "&" & sLabel2 & "=" & (nRawID + 1)
       End Select
    
    '------------------> Return data element
    smfGetADVFNElement = RCHGetTableCell(sURL & nPageId, nColumn, pFind1, pFind2, , , , "</TABLE", , vError)
    
    Exit Function

ErrorExit: smfGetADVFNElement = vError
                   
   End Function
Sub smfSetAdvFNPrefix(p1 As String)
    sAdvFNPrefix = p1
    End Sub

Sub LoadAdvFNPrefix()
    sAdvFNPrefix = "www"
    On Error GoTo ErrorExit
    Open ThisWorkbook.Path & "\smf-AdvFN-Prefix.txt" For Input As #1
    Line Input #1, sAdvFNPrefix
    Close #1
ErrorExit:
    End Sub



#If Mac Then
Private Function smfMacADVFNLatestYearEnd(ByVal pTicker As String, ByVal pFind2 As String, _
                                         ByVal pError As Variant, ByVal pType As Integer) As Variant
    Dim page As String, url As String, ticker As String, section As Variant
    Dim header As String, cellText As String, firstText As String, lastText As String
    Dim cellAt As Long, contentAt As Long, closeAt As Long, cellCount As Long
    Dim fiscalDate As Date
    On Error GoTo Failed
    smfMacADVFNLatestYearEnd = pError
    ticker = smfConvertYahooTicker(pTicker, "ADVFN")
    If Len(ticker) = 0 Then Exit Function
    If sAdvFNPrefix = "" Then LoadAdvFNPrefix
    url = "https://" & sAdvFNPrefix & ".advfn.com/p.php?pid=financials&symbol=" & ticker
    page = smfGetWebPage(url, pType, 0)
    For Each section In Array("annualIncomeStatement", "annualBalanceSheet", "annualCashFlow")
        header = smfMacADVFNAnnualHeader(page, CStr(section))
        If Len(header) > 0 Then
            ' Keep the caller's additional search within the selected annual header.
            If pFind2 > " " Then
                If InStr(1, header, pFind2, vbTextCompare) = 0 Then GoTo NextSection
            End If
            ' The shared page cache converts TH tags to TD tags. Accept either form.
            header = Replace(header, "<th", "<td", 1, -1, vbTextCompare)
            header = Replace(header, "</th", "</td", 1, -1, vbTextCompare)
            cellAt = 1
            cellCount = 0
            firstText = ""
            lastText = ""
            Do
                cellAt = InStr(cellAt, header, "<td", vbTextCompare)
                If cellAt = 0 Then Exit Do
                contentAt = InStr(cellAt, header, ">")
                closeAt = InStr(contentAt + 1, header, "</td", vbTextCompare)
                If contentAt = 0 Or closeAt = 0 Then Exit Function
                cellText = smfMacADVFNText(smfMacStripHTML(Mid$(header, contentAt + 1, closeAt - contentAt - 1)))
                cellCount = cellCount + 1
                If cellCount = 1 Then firstText = cellText
                If Len(cellText) > 0 Then lastText = cellText
                cellAt = closeAt + 4
            Loop
            If cellCount > 1 And smfMacADVFNParseYearEnd(firstText, fiscalDate) Then
                If Len(lastText) = 4 And Not lastText Like "*[!0-9]*" Then
                    If CLng(lastText) = Year(fiscalDate) Then
                        smfMacADVFNLatestYearEnd = fiscalDate
                        Exit Function
                    End If
                End If
            End If
        End If
NextSection:
    Next section
Failed:
End Function

Private Function smfMacADVFNAnnualHeader(ByVal pHTML As String, ByVal pSection As String) As String
    Dim markerAt As Long, divAt As Long, openAt As Long, pos As Long, tagAt As Long
    Dim tagEnd As Long, depth As Long, sectionEnd As Long, headerAt As Long, headerEnd As Long
    Dim tag As String, sectionHTML As String, headerTag As String
    ' Restrict the lookup to an actual annual container, excluding navigation/JavaScript names.
    markerAt = InStr(1, pHTML, "id=""" & pSection & """", vbTextCompare)
    If markerAt = 0 Then markerAt = InStr(1, pHTML, "id='" & pSection & "'", vbTextCompare)
    If markerAt = 0 Then Exit Function
    divAt = InStrRev(pHTML, "<div", markerAt, vbTextCompare)
    If divAt = 0 Then Exit Function
    openAt = InStr(divAt, pHTML, ">")
    If openAt = 0 Or openAt < markerAt Then Exit Function
    pos = openAt + 1
    depth = 1
    Do While depth > 0
        tagAt = InStr(pos, pHTML, "<")
        If tagAt = 0 Then Exit Function
        If Mid$(pHTML, tagAt, 4) = "<!--" Then
            tagEnd = InStr(tagAt + 4, pHTML, "-->")
            If tagEnd = 0 Then Exit Function
            pos = tagEnd + 3
        Else
            tagEnd = InStr(tagAt, pHTML, ">")
            If tagEnd = 0 Then Exit Function
            tag = LCase$(smfMacADVFNText(Mid$(pHTML, tagAt + 1, tagEnd - tagAt - 1)))
            If tag = "div" Or Left$(tag, 4) = "div " Then depth = depth + 1
            If tag = "/div" Then depth = depth - 1
            If depth = 0 Then sectionEnd = tagAt
            pos = tagEnd + 1
        End If
    Loop
    sectionHTML = Mid$(pHTML, openAt + 1, sectionEnd - openAt - 1)
    headerTag = "thead"
    headerAt = InStr(1, sectionHTML, "<" & headerTag, vbTextCompare)
    If headerAt = 0 Then
        ' The normal cache conversion also changes THEAD to TDEAD.
        headerTag = "tdead"
        headerAt = InStr(1, sectionHTML, "<" & headerTag, vbTextCompare)
    End If
    If headerAt = 0 Then Exit Function
    headerAt = InStr(headerAt, sectionHTML, ">")
    headerEnd = InStr(headerAt + 1, sectionHTML, "</" & headerTag, vbTextCompare)
    If headerAt = 0 Or headerEnd = 0 Then Exit Function
    smfMacADVFNAnnualHeader = Mid$(sectionHTML, headerAt + 1, headerEnd - headerAt - 1)
End Function

Private Function smfMacADVFNParseYearEnd(ByVal pText As String, ByRef pDate As Date) As Boolean
    Dim text As String, parts As Variant, months As Variant, monthNumber As Long
    Dim dayNumber As Long, yearNumber As Long, i As Long, candidate As Date
    On Error GoTo Failed
    text = LCase$(smfMacADVFNText(pText))
    Select Case True
        Case Left$(text, 14) = "year end date ": text = Mid$(text, 15)
        Case Left$(text, 9) = "year end ": text = Mid$(text, 10)
        Case Else: Exit Function
    End Select
    text = smfMacADVFNText(Replace(text, ",", " "))
    parts = Split(text, " ")
    If UBound(parts) <> 2 Then Exit Function
    months = Split("january february march april may june july august september october november december", " ")
    For i = 0 To UBound(months)
        If parts(0) = months(i) Then monthNumber = i + 1
    Next i
    If monthNumber = 0 Then Exit Function
    If Len(parts(1)) < 1 Or Len(parts(1)) > 2 Or parts(1) Like "*[!0-9]*" Then Exit Function
    If Len(parts(2)) <> 4 Or parts(2) Like "*[!0-9]*" Then Exit Function
    dayNumber = CLng(parts(1))
    yearNumber = CLng(parts(2))
    If dayNumber < 1 Or dayNumber > 31 Or yearNumber < 100 Then Exit Function
    candidate = DateSerial(yearNumber, monthNumber, dayNumber)
    ' DateSerial rolls invalid days into the next month; reject that rollover.
    If Year(candidate) <> yearNumber Or Month(candidate) <> monthNumber Or Day(candidate) <> dayNumber Then Exit Function
    pDate = candidate
    smfMacADVFNParseYearEnd = True
Failed:
End Function

Private Function smfMacADVFNText(ByVal pText As String) As String
    Dim text As String
    text = Replace(Replace(Replace(pText, vbCr, " "), vbLf, " "), vbTab, " ")
    text = Replace(text, ChrW(160), " ")
    Do While InStr(text, "  ") > 0
        text = Replace(text, "  ", " ")
    Loop
    smfMacADVFNText = Trim$(text)
End Function
#End If

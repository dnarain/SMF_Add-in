Attribute VB_Name = "smfGetYahooHistoryCSV_"
Option Explicit

Function smfGetYahooHistoryCSV(ByVal pTicker As String, _
                   Optional ByVal pStartDate As Variant = "", _
                   Optional ByVal pEndDate As Variant = "", _
                   Optional ByVal pPeriod As String = "d", _
                   Optional ByVal pRows As Integer = 10000, _
                   Optional ByVal pCols As Integer = 7)

    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to process CSV file from Yahoo
    '-----------------------------------------------------------------------------------------------------------*
    ' 2017.05.31 -- Added
    '-----------------------------------------------------------------------------------------------------------*
                   
    ReDim vData(1 To 1, 1 To 1) As Variant
    vData(1, 1) = "Error"
    
    On Error GoTo ErrorExit
    
    '------------------> Null Return Item
    If pTicker = "None" Or pTicker = "" Then
       vData(1, 1) = "None"
       GoTo ErrorExit
       End If
       
    '------------------> Verify and Process starting and ending dates
    Dim dBegin As Double, dEnd As Double
    vData(1, 1) = "Error on starting date: " & pStartDate
    Select Case True
          Case VarType(pStartDate) = vbDate Or VarType(pStartDate) = vbDouble
               dBegin = smfDate2Unix(pStartDate)
          Case pStartDate = ""
               dBegin = smfDate2Unix(DateValue("1/1/1970"))
          Case Else
               dBegin = smfDate2Unix(DateValue(pStartDate))
          End Select
    vData(1, 1) = "Error on ending date: " & pEndDate
    Select Case True
          Case VarType(pEndDate) = vbDate Or VarType(pEndDate) = vbDouble
               dEnd = smfDate2Unix(Int(pEndDate) + 1)
          Case pEndDate = ""
               dEnd = smfDate2Unix(Int(Now) + 1)
          Case Else
               dEnd = smfDate2Unix(Int(DateValue(pEndDate)) + 1)
          End Select
     If dBegin > dEnd Then
        vData(1, 1) = "Error: Starting date cannot be after ending date: " & pStartDate & "," & pEndDate
        GoTo ErrorExit
        End If
       
    '------------------> Process period
    Dim sPeriod As String, sEvent As String, sInterval As String
    sPeriod = UCase(pPeriod)
    Select Case sPeriod
       Case "D": sEvent = "history": sInterval = "1d"
       Case "W": sEvent = "history": sInterval = "1wk"
       Case "M": sEvent = "history": sInterval = "1mo"
       Case "S": sEvent = "split": sInterval = "1d"
       Case "V": sEvent = "div": sInterval = "1d"
       Case Else
            vData(1, 1) = "Error on period: " & pPeriod
            GoTo ErrorExit
       End Select
    
    '------------------> Determine size of array to return and initialize array
    Dim iRows As Integer, iCols As Integer, i1 As Integer, i2 As Integer
    iRows = pRows  ' Rows
    iCols = pCols  ' Columns
    On Error Resume Next
    iRows = Application.Caller.Rows.Count
    iCols = Application.Caller.Columns.Count
    On Error GoTo ErrorExit
  
    ReDim vData(1 To iRows, 1 To iCols) As Variant
    For i1 = 1 To iRows
        For i2 = 1 To iCols
            vData(i1, i2) = ""
            Next i2
        Next i1
    
    '------------------> Get CSV file
    Dim sURL As String
    sURL = "https://query1.finance.yahoo.com/v7/finance/download/" & UCase(pTicker) & _
           "?period1=" & dBegin & "&period2=" & dEnd & "&interval=" & sInterval & "&events=" & sEvent
'           "?period1=" & dBegin & "&period2=" & dEnd & "&interval=" & sInterval & "&events=" & sEvent & "&crumb="

    vData = smfGetCSVFile(sURL, ",", iRows, iCols)

ErrorExit:
    smfGetYahooHistoryCSV = vData
                   
   End Function



Function smfGetYahooHistoryCSVData(Optional ByRef pURL As String = "https://query1.finance.yahoo.com/v7/finance/download/MMM?period1=1493610466&period2=1496202466&interval=1d&events=history&crumb=")

   '-----------------------------------------------------------------------------------------------------------*
   ' User defined function to download historical quotes data file from Yahoo
   '-----------------------------------------------------------------------------------------------------------*
   ' 2017.05.31 -- Adapted from code at http://www.xlautomation.com.au/free-spreadsheets/yahoo-historical-price-extract
   '-----------------------------------------------------------------------------------------------------------*

   Dim sURL As String
   Dim sCrumb As String, sCookie As String, sResult As String
   Dim i1 As Integer, b1 As Boolean
   Dim oData As Object
   
#If Mac Then
   If smfMacYahooDownloadURL(pURL) Then
      smfGetYahooHistoryCSVData = smfMacYahooDownloadCSV(pURL)
   Else
      smfGetYahooHistoryCSVData = smfMacFetch(pURL, "YAHOO")
   End If
#Else
   sURL = "https://finance.yahoo.com/lookup?s=%7B0%7D"
   Set oData = New WinHttp.WinHttpRequest
   b1 = True
   For i1 = 1 To 5
       With oData
            .Open "GET", sURL, b1
            Select Case i1
               Case 1: .setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
               Case Else: .setRequestHeader "Cookie", sCookie
               End Select
            .send
            .waitForResponse
            Select Case i1
               Case 1
                    sCrumb = smfStrExtr(smfStrExtr(.responseText, "CrumbStore", "~"), """:""", """")
                    sCookie = Split(.getResponseHeader("Set-Cookie"), ";")(0)
                    sURL = pURL + sCrumb
               Case Else
                    sResult = .responseText
                    If Left(sResult, 4) = "Date" Then Exit For
               End Select
            End With
       b1 = False
       Next i1
   
    smfGetYahooHistoryCSVData = sResult
#End If
      
End Function

#If Mac Then
' Keep the legacy CSV result while retrieving Yahoo's current chart response.
' Only the two Yahoo download endpoints qualify for this conversion.
Public Function smfMacYahooDownloadURL(ByVal pURL As String) As Boolean
    Dim value As String, prefix As Variant
    value = LCase$(pURL)
    If InStr(value, vbCr) > 0 Or InStr(value, vbLf) > 0 Then Exit Function
    For Each prefix In Array("https://query1.finance.yahoo.com/v7/finance/download/", _
                             "https://query2.finance.yahoo.com/v7/finance/download/")
        If Left$(value, Len(prefix)) = prefix Then
            smfMacYahooDownloadURL = True
            Exit Function
        End If
    Next prefix
End Function

Public Function smfMacYahooDownloadCSV(ByVal pURL As String) As String
    Dim question As Long, tail As String, symbol As String, query As String
    Dim eventName As String, interval As String, first As String, last As String
    Dim part As Variant, pair As Variant, chartURL As String, payload As String
    On Error GoTo Failed
    If Not smfMacYahooDownloadURL(pURL) Then Err.Raise 5, , "Expected a Yahoo download URL"
    tail = Mid$(pURL, InStr(1, pURL, "/download/", vbTextCompare) + 10)
    question = InStr(tail, "?")
    If question < 2 Then Err.Raise 5, , "Missing ticker or history parameters"
    symbol = Left$(tail, question - 1)
    If InStr(symbol, "/") > 0 Or InStr(symbol, "#") > 0 Then Err.Raise 5, , "Invalid ticker path"
    query = Mid$(tail, question + 1)
    eventName = "history"
    interval = "1d"
    For Each part In Split(query, "&")
        pair = Split(CStr(part), "=", 2)
        If UBound(pair) = 1 Then
            Select Case LCase$(pair(0))
                Case "period1": first = pair(1)
                Case "period2": last = pair(1)
                Case "interval": interval = pair(1)
                Case "events": eventName = LCase$(pair(1))
            End Select
        End If
    Next part
    If Not MacHistoryInteger(first) Or Not MacHistoryInteger(last) Then Err.Raise 5, , "Invalid history dates"
    If CDbl(first) >= CDbl(last) Then Err.Raise 5, , "Starting date must precede ending date"
    If interval <> "1d" And interval <> "1wk" And interval <> "1mo" Then Err.Raise 5, , "Unsupported history interval"
    If eventName <> "history" And eventName <> "div" And eventName <> "split" Then Err.Raise 5, , "Unsupported history event"
    chartURL = "https://query1.finance.yahoo.com/v8/finance/chart/" & symbol & _
        "?period1=" & first & "&period2=" & last & "&interval=" & interval & _
        "&events=div%2Csplits&includeAdjustedClose=true"
    payload = smfGetWebPage(chartURL, 0, 0)
    If Left$(payload, 5) = "Error" Then
        smfMacYahooDownloadCSV = payload
        Exit Function
    End If
    smfMacYahooDownloadCSV = smfMacYahooChartCSV(payload, eventName)
    Exit Function
Failed:
    smfMacYahooDownloadCSV = "Error -- Yahoo history: " & Err.Description
End Function

' Public so fixtures can check returned dates, OHLC, adjusted close and events.
Public Function smfMacYahooChartCSV(ByVal pJSON As String, Optional ByVal pEvent As String = "history") As String
    Dim root As Object, chart As Object, data As Object, quote As Object, adjusted As Object
    Dim stamps As Object, events As Object, record As Object, keys As Variant
    Dim lines() As String, offset As Double, i As Long, j As Long, key As Variant, temp As Variant
    Dim fields As Variant, f As Variant, value As Variant, stamp As Double, line As String
    On Error GoTo Failed
    Set root = JsonConverter.ParseJson(pJSON)
    Set chart = root("chart")
    If Not IsNull(chart("error")) Then Err.Raise 5, , "The provider returned a chart error"
    If chart("result").Count <> 1 Then Err.Raise 5, , "Missing chart result"
    Set data = chart("result")(1)
    offset = CDbl(data("meta")("gmtoffset"))
    Select Case LCase$(pEvent)
        Case "history"
            Set stamps = data("timestamp")
            Set quote = data("indicators")("quote")(1)
            Set adjusted = data("indicators")("adjclose")(1)("adjclose")
            fields = Array("open", "high", "low", "close", "volume")
            If adjusted.Count <> stamps.Count Then Err.Raise 5, , "Adjusted-close dates do not match"
            For Each f In fields
                If quote(CStr(f)).Count <> stamps.Count Then Err.Raise 5, , "Quote dates do not match"
            Next f
            ReDim lines(0 To stamps.Count) As String
            lines(0) = "Date,Open,High,Low,Close,Adj Close,Volume"
            For i = 1 To stamps.Count
                stamp = CDbl(stamps(i))
                line = MacHistoryDate(stamp, offset)
                For Each f In Array("open", "high", "low", "close")
                    line = line & "," & MacHistoryNumber(quote(CStr(f))(i))
                Next f
                lines(i) = line & "," & MacHistoryNumber(adjusted(i)) & "," & MacHistoryNumber(quote("volume")(i))
            Next i
        Case "div", "split"
            ReDim lines(0 To 0) As String
            If LCase$(pEvent) = "div" Then
                lines(0) = "Date,Dividends"
                key = "dividends"
            Else
                lines(0) = "Date,Stock Splits"
                key = "splits"
            End If
            If data.Exists("events") Then
                If data("events").Exists(CStr(key)) Then
                    Set events = data("events")(CStr(key))
                    keys = events.Keys
                    If events.Count > 0 Then
                        For i = LBound(keys) To UBound(keys) - 1
                            For j = i + 1 To UBound(keys)
                                If CDbl(keys(i)) > CDbl(keys(j)) Then
                                    temp = keys(i): keys(i) = keys(j): keys(j) = temp
                                End If
                            Next j
                        Next i
                        ReDim Preserve lines(0 To events.Count) As String
                        For i = LBound(keys) To UBound(keys)
                            Set record = events(keys(i))
                            line = MacHistoryDate(CDbl(record("date")), offset)
                            If LCase$(pEvent) = "div" Then
                                value = MacHistoryNumber(record("amount"))
                            Else
                                value = MacHistoryNumber(record("numerator")) & ":" & MacHistoryNumber(record("denominator"))
                            End If
                            lines(i - LBound(keys) + 1) = line & "," & CStr(value)
                        Next i
                    End If
                End If
            End If
        Case Else
            Err.Raise 5, , "Unsupported history event"
    End Select
    smfMacYahooChartCSV = Join(lines, vbLf) & vbLf
    Exit Function
Failed:
    smfMacYahooChartCSV = "Error -- Yahoo history response: " & Err.Description
End Function

Public Function SMFMacHistoryCSVSelfTest() As Long
    Dim fixture As String, csv As String, expected As String, count As Long
    fixture = "{""chart"":{""error"":null,""result"":[{""meta"":{""gmtoffset"":-18000}," & _
        """timestamp"":[1793629800,1793716200],""indicators"":{""quote"":[{""open"":[100,110]," & _
        """high"":[105,115],""low"":[95,105],""close"":[101,111],""volume"":[1234,2345]}]," & _
        """adjclose"":[{""adjclose"":[50.5,null]}]},""events"":{""dividends"":{" & _
        """1793716200"":{""date"":1793716200,""amount"":0.5}," & _
        """1793629800"":{""date"":1793629800,""amount"":0.25}},""splits"":{" & _
        """1793629800"":{""date"":1793629800,""numerator"":4,""denominator"":1}}}}]}}"
    csv = smfMacYahooChartCSV(fixture)
    expected = "Date,Open,High,Low,Close,Adj Close,Volume" & vbLf & _
        "2026-11-02,100,105,95,101,50.5,1234" & vbLf & _
        "2026-11-03,110,115,105,111,null,2345" & vbLf
    If csv <> expected Then Err.Raise 5, , "History OHLC, adjusted close, null, dates or volume changed: " & csv
    count = count + 1
    csv = smfMacYahooChartCSV(fixture, "div")
    expected = "Date,Dividends" & vbLf & "2026-11-02,0.25" & vbLf & "2026-11-03,0.5" & vbLf
    If csv <> expected Then Err.Raise 5, , "Dividend ordering, amounts or dates changed"
    count = count + 1
    csv = smfMacYahooChartCSV(fixture, "split")
    expected = "Date,Stock Splits" & vbLf & "2026-11-02,4:1" & vbLf
    If csv <> expected Then Err.Raise 5, , "Split ratio or date changed"
    count = count + 1
    csv = smfMacYahooChartCSV(Replace(fixture, """open"":[100,110]", """open"":[100]"))
    If Left$(csv, 5) <> "Error" Then Err.Raise 5, , "Mismatched quote dates were accepted"
    count = count + 1
    csv = smfMacYahooChartCSV("{""chart"":{""error"":{""code"":""Bad Request""},""result"":null}}")
    If Left$(csv, 5) <> "Error" Then Err.Raise 5, , "Provider error became data"
    count = count + 1
    csv = smfMacYahooChartCSV(Replace(fixture, """dividends"":{", """otherEvent"":{"), "div")
    If csv <> "Date,Dividends" & vbLf Then Err.Raise 5, , "Missing dividends produced observations"
    count = count + 1
    If Not smfMacYahooDownloadURL("https://query1.finance.yahoo.com/v7/finance/download/AAPL?period1=1") Then Err.Raise 5
    count = count + 1
    If Not smfMacYahooDownloadURL("https://query2.finance.yahoo.com/v7/finance/download/IBM?period1=1") Then Err.Raise 5
    count = count + 1
    If smfMacYahooDownloadURL("https://example.invalid/v7/finance/download/AAPL") Then Err.Raise 5
    count = count + 1
    If smfMacYahooDownloadURL("https://query1.finance.yahoo.com.example.invalid/v7/finance/download/AAPL") Then Err.Raise 5
    count = count + 1
    csv = smfMacYahooDownloadCSV("https://query1.finance.yahoo.com/v7/finance/download/AAPL?period1=1x&period2=20")
    If Left$(csv, 5) <> "Error" Then Err.Raise 5, , "Invalid dates reached the network"
    count = count + 1
    csv = smfMacYahooDownloadCSV("https://query1.finance.yahoo.com/v7/finance/download/AAPL?period1=20&period2=1")
    If Left$(csv, 5) <> "Error" Then Err.Raise 5, , "Reversed dates reached the network"
    count = count + 1
    SMFMacHistoryCSVSelfTest = count
End Function

Private Function MacHistoryInteger(ByVal value As String) As Boolean
    Dim i As Long, first As Long
    first = 1
    If Left$(value, 1) = "-" Then first = 2
    If Len(value) < first Or Len(value) > 13 Then Exit Function
    For i = first To Len(value)
        If Mid$(value, i, 1) < "0" Or Mid$(value, i, 1) > "9" Then Exit Function
    Next i
    MacHistoryInteger = True
End Function

Private Function MacHistoryDate(ByVal stamp As Double, ByVal offset As Double) As String
    MacHistoryDate = Format$(DateSerial(1970, 1, 1) + (stamp + offset) / 86400#, "yyyy-mm-dd")
End Function

Private Function MacHistoryNumber(ByVal value As Variant) As String
    If IsNull(value) Then
        MacHistoryNumber = "null"
    ElseIf VarType(value) = vbString Then
        ' VBA-JSON preserves numbers longer than15digits as strings.
        ' Keep their invariant decimal text, without accepting CSV separators.
        If Not MacHistoryJSONNumber(CStr(value)) Then Err.Raise 5, , "Expected a numeric quote value"
        MacHistoryNumber = CStr(value)
    ElseIf IsNumeric(value) And VarType(value) <> vbBoolean And VarType(value) <> vbString Then
        MacHistoryNumber = Trim$(Str$(CDbl(value)))
        If Left$(MacHistoryNumber, 1) = "." Then MacHistoryNumber = "0" & MacHistoryNumber
        If Left$(MacHistoryNumber, 2) = "-." Then MacHistoryNumber = "-0" & Mid$(MacHistoryNumber, 2)
    Else
        Err.Raise 5, , "Expected a numeric quote value"
    End If
End Function

Private Function MacHistoryJSONNumber(ByVal value As String) As Boolean
    Dim position As Long, firstDigit As Long, size As Long, character As String
    size = Len(value)
    If size = 0 Or size > 128 Then Exit Function
    position = 1
    If Left$(value, 1) = "-" Then position = position + 1
    character = Mid$(value, position, 1)
    If character = "0" Then
        position = position + 1
    ElseIf character >= "1" And character <= "9" And Len(character) = 1 Then
        Do While position <= size
            character = Mid$(value, position, 1)
            If character < "0" Or character > "9" Then Exit Do
            position = position + 1
        Loop
    Else
        Exit Function
    End If
    If Mid$(value, position, 1) = "." Then
        position = position + 1: firstDigit = position
        Do While position <= size
            character = Mid$(value, position, 1)
            If character < "0" Or character > "9" Then Exit Do
            position = position + 1
        Loop
        If position = firstDigit Then Exit Function
    End If
    character = Mid$(value, position, 1)
    If character = "e" Or character = "E" Then
        position = position + 1
        character = Mid$(value, position, 1)
        If character = "+" Or character = "-" Then position = position + 1
        firstDigit = position
        Do While position <= size
            character = Mid$(value, position, 1)
            If character < "0" Or character > "9" Then Exit Do
            position = position + 1
        Loop
        If position = firstDigit Then Exit Function
    End If
    MacHistoryJSONNumber = (position = size + 1)
End Function
#End If

Attribute VB_Name = "modRSITargetPrices"
Option Explicit
Public Function smfRSITargetPrices(ByVal pTicker As String, _
                          Optional ByVal pLoTrigger As Integer = 20, _
                          Optional ByVal pHiTrigger As Integer = 80, _
                          Optional ByVal pItems As Variant = "010203040506070809101112")
#If Mac Then
    smfRSITargetPrices = smfMacRSITargetPrices(pTicker, pLoTrigger, pHiTrigger, pItems)
#Else
                        
   '-----------------------------------------------------------------------------------------------------------*
   ' Function to return RSI indicator buy and sell target prices
   '-----------------------------------------------------------------------------------------------------------*
   ' 2012.01.06 -- Created by Randy Harmelink (rharmelink@gmail.com)
   ' 2017.05.19 -- Change to use smfGetYahooHistory()
   ' 2017.05.23 -- Create starting date for smfGetYahooHistory() call
   '-----------------------------------------------------------------------------------------------------------*
   ' Samples of use:
   '
   '    =smfRSITargetPrices("MMM",20,80)
   '    =smfRSITargetPrices("MMM",20,80,"010203")
   '-----------------------------------------------------------------------------------------------------------*
     
   Const kItems = 12
   Dim vData(1 To 1, 1 To kItems) As Variant
   
   On Error GoTo ErrorExit
   vData(1, 1) = "Error"
   
   Dim i1 As Integer
   If pTicker = "Header" Or pTicker = "Ticker" Or pTicker = "Symbol" Then
      For i1 = 1 To kItems
          Select Case Mid(pItems, 2 * i1 - 1, 2)
             Case "01": vData(1, i1) = "Current RSI"
             Case "02": vData(1, i1) = "Buy Target Price"
             Case "03": vData(1, i1) = "Sell Target Price"
             Case "04": vData(1, i1) = "Last Traded Price"
             Case "05": vData(1, i1) = "Bid Price"
             Case "06": vData(1, i1) = "Ask Price"
             Case "07": vData(1, i1) = "Open Price"
             Case "08": vData(1, i1) = "Low Price"
             Case "09": vData(1, i1) = "High Price"
             Case "10": vData(1, i1) = "Volume"
             Case "11": vData(1, i1) = "Previous Close"
             Case "12": vData(1, i1) = "Previous RSI"
             Case Else: vData(1, i1) = "--"
             End Select
          Next i1
      GoTo ErrorExit
      End If
   
   Dim vCQ As Variant, vHQ As Variant, vRSI As Variant
   vCQ = RCHGetYahooQuotes(pTicker, "l1baoghvd1t1")
   'vHQ = RCHGetYahooHistory2(pTicker, , , , , , , , , , 1, 1, 22, 6)
   vHQ = smfGetYahooHistory(pTicker, Int(Now) - 40, , , , , 1, 22, 6)
   vRSI = SMFTech(vHQ, "RSI", 2)
   
   For i1 = 1 To kItems
       Select Case Mid(pItems, 2 * i1 - 1, 2)
          Case "01": vData(1, i1) = 100 - 100 / (1 + (vRSI(22, 5) + Application.WorksheetFunction.Max(vCQ(1, 1) - vHQ(22, 5), 0)) / (vRSI(22, 6) + Application.WorksheetFunction.Max(0, vHQ(22, 5) - vCQ(1, 1))))
          Case "02": vData(1, i1) = IIf(pLoTrigger > vRSI(22, 1), "--", vHQ(22, 5) + vRSI(22, 6) - vRSI(22, 5) * (100 - pLoTrigger) / pLoTrigger)
          Case "03": vData(1, i1) = IIf(pHiTrigger < vRSI(22, 1), "--", vHQ(22, 5) - vRSI(22, 5) + vRSI(22, 6) * pHiTrigger / (100 - pHiTrigger))
          Case "04": vData(1, i1) = vCQ(1, 1)
          Case "05": vData(1, i1) = vCQ(1, 2)
          Case "06": vData(1, i1) = vCQ(1, 3)
          Case "07": vData(1, i1) = vCQ(1, 4)
          Case "08": vData(1, i1) = vCQ(1, 5)
          Case "09": vData(1, i1) = vCQ(1, 6)
          Case "10": vData(1, i1) = vCQ(1, 7)
          Case "11": vData(1, i1) = vHQ(22, 5)
          Case "12": vData(1, i1) = vRSI(22, 1)
          Case Else: vData(1, i1) = "--"
          End Select
       Next i1
  
ErrorExit:
   
   smfRSITargetPrices = vData
   
#End If
   End Function

#If Mac Then
' RSI(2) uses 21 prior completed observations, matching the old 22-row/header input.
' The latest regular quote defines the session; its daily bar is excluded.
Private Function smfMacRSITargetPrices(ByVal ticker As String, ByVal lowTrigger As Integer, _
        ByVal highTrigger As Integer, ByVal requestedItems As Variant) As Variant
    Dim output(1 To 1, 1 To 12) As Variant, values(1 To 12) As Variant
    Dim selected(1 To 12) As Long, labels As Variant, items As String, itemValue As Variant
    Dim i As Long, item As Long, hasItems As Boolean
    Dim needRSI As Boolean, needPrevious As Boolean, needBar As Boolean
    Dim needBidAsk As Boolean, needLow As Boolean, needHigh As Boolean
    Dim root As Object, snapshot As Object, meta As Object, bars As Object
    Dim payload As String, encoded As String, quoteDate As Date, bar As Long
    Dim currentPrice As Double, quoteTime As Double, offset As Double
    Dim previousClose As Double, averageGain As Double, averageLoss As Double
    Dim previousRSI As Double, change As Double, gain As Double, loss As Double
    Dim candidate As Double, requiredCloses As Long
    On Error GoTo Failed
    labels = Split("Current RSI|Buy Target Price|Sell Target Price|Last Traded Price|Bid Price|Ask Price|Open Price|Low Price|High Price|Volume|Previous Close|Previous RSI", "|")
    If IsObject(requestedItems) Then
        If TypeName(requestedItems) <> "Range" Then GoTo Failed
        If requestedItems.Cells.Count <> 1 Then GoTo Failed
        itemValue = requestedItems.Cells(1, 1).Value2
    Else
        itemValue = requestedItems
    End If
    If IsError(itemValue) Or IsNull(itemValue) Then GoTo Failed
    If IsArray(itemValue) Then GoTo Failed
    items = CStr(itemValue)
    ticker = UCase$(Trim$(ticker))
    For i = 1 To 12
        output(1, i) = "--"
        values(i) = "--"
        item = smfMacRSIItemIndex(Mid$(items, 2 * i - 1, 2))
        selected(i) = item
        If item <> 0 Then hasItems = True
        Select Case item
            Case 1, 2, 3, 12: needRSI = True
            Case 11: needPrevious = True
            Case 7 To 10: needBar = True
        End Select
        If item = 2 Then needLow = True
        If item = 3 Then needHigh = True
        If item = 5 Or item = 6 Then needBidAsk = True
    Next i
    If ticker = "HEADER" Or ticker = "TICKER" Or ticker = "SYMBOL" Then
        For i = 1 To 12
            If selected(i) <> 0 Then output(1, i) = labels(selected(i) - 1)
        Next i
        smfMacRSITargetPrices = output
        Exit Function
    End If
    If ticker = "NONE" Or ticker = "" Then
        Erase output
        output(1, 1) = "None"
        smfMacRSITargetPrices = output
        Exit Function
    End If
    If Not hasItems Then
        smfMacRSITargetPrices = output
        Exit Function
    End If
    If needLow Then
        If lowTrigger <= 0 Or lowTrigger >= 100 Then GoTo Failed
    End If
    If needHigh Then
        If highTrigger <= 0 Or highTrigger >= 100 Then GoTo Failed
    End If
    If needLow And needHigh Then
        If lowTrigger >= highTrigger Then GoTo Failed
    End If
    encoded = smfMacRSITicker(ticker)
    If encoded = "" Then GoTo Failed
    payload = smfGetWebPage("https://query1.finance.yahoo.com/v8/finance/chart/" & encoded & _
        "?range=3mo&interval=1d&events=div%2Csplits&includeAdjustedClose=true", 0, 0)
    Set root = smfMacRSIParseJSON(payload)
    If Not IsNull(root("chart")("error")) Then GoTo Failed
    If root("chart")("result").Count <> 1 Then GoTo Failed
    Set snapshot = root("chart")("result")(1)
    Set meta = snapshot("meta")
    If UCase$(CStr(meta("symbol"))) <> ticker Then GoTo Failed
    If Not smfMacRSINumber(meta("regularMarketPrice"), currentPrice) Then GoTo Failed
    If currentPrice <= 0 Then GoTo Failed
    If Not smfMacRSINumber(meta("regularMarketTime"), quoteTime) Then GoTo Failed
    If Not smfMacRSINumber(meta("gmtoffset"), offset) Then GoTo Failed
    If Abs(offset) > 86400 Then GoTo Failed
    quoteDate = smfMacRSIDate(quoteTime, offset)
    If quoteDate = 0 Then GoTo Failed
    values(4) = currentPrice
    If needRSI Then
        requiredCloses = 21
    ElseIf needPrevious Then
        requiredCloses = 1
    End If
    If requiredCloses > 0 Or needBar Then
        If Not smfMacRSIHistory(snapshot, quoteDate, offset, requiredCloses, previousClose, _
                    averageGain, averageLoss, bars, bar) Then
            If requiredCloses > 0 Then GoTo Failed
        End If
    End If
    If requiredCloses > 0 Then values(11) = previousClose
    If needRSI Then
        previousRSI = smfMacRSIRatio(averageGain, averageLoss)
        change = currentPrice - previousClose
        gain = averageGain
        loss = averageLoss
        If change > 0 Then gain = gain + change
        If change < 0 Then loss = loss - change
        values(1) = smfMacRSIRatio(gain / 2, loss / 2)
        values(12) = previousRSI
        ' A flat prior series has no unique finite price for an interior RSI.
        If averageGain + averageLoss > 0 Then
            If needLow Then
                If lowTrigger <= previousRSI Then
                    candidate = previousClose + averageLoss - averageGain * (100 - lowTrigger) / lowTrigger
                    If candidate > 0 Then values(2) = candidate
                End If
            End If
            If needHigh Then
                If highTrigger >= previousRSI Then
                    candidate = previousClose - averageGain + averageLoss * highTrigger / (100 - highTrigger)
                    If candidate > 0 Then values(3) = candidate
                End If
            End If
        End If
    End If
    values(7) = smfMacRSIBarNumber(bars, "open", bar, True)
    values(8) = smfMacRSIOptionalNumber(meta, "regularMarketDayLow", True)
    values(9) = smfMacRSIOptionalNumber(meta, "regularMarketDayHigh", True)
    values(10) = smfMacRSIOptionalNumber(meta, "regularMarketVolume", False)
    If VarType(values(8)) = vbString Then values(8) = smfMacRSIBarNumber(bars, "low", bar, True)
    If VarType(values(9)) = vbString Then values(9) = smfMacRSIBarNumber(bars, "high", bar, True)
    If VarType(values(10)) = vbString Then values(10) = smfMacRSIBarNumber(bars, "volume", bar, False)
    If needBidAsk Then smfMacRSIBidAsk ticker, encoded, quoteDate, offset, values(5), values(6)
    For i = 1 To 12
        If selected(i) <> 0 Then output(1, i) = values(selected(i))
    Next i
    smfMacRSITargetPrices = output
    Exit Function
Failed:
    Erase output
    output(1, 1) = "Error"
    smfMacRSITargetPrices = output
End Function

Private Function smfMacRSIParseJSON(ByVal payload As String) As Object
    Dim savedLarge As Boolean, failureNumber As Long
    Dim failureSource As String, failureText As String
    savedLarge = JsonConverter.JsonOptions.UseDoubleForLargeNumbers
    On Error GoTo Failed
    ' Yahoo prices/closes can have more than 15 significant digits.
    ' Only numeric JSON tokens become Double; quoted strings stay strings.
    JsonConverter.JsonOptions.UseDoubleForLargeNumbers = True
    Set smfMacRSIParseJSON = JsonConverter.ParseJson(payload)
    JsonConverter.JsonOptions.UseDoubleForLargeNumbers = savedLarge
    Exit Function
Failed:
    failureNumber = Err.Number
    failureSource = Err.Source
    failureText = Err.Description
    JsonConverter.JsonOptions.UseDoubleForLargeNumbers = savedLarge
    Err.Raise failureNumber, failureSource, failureText
End Function

Private Function smfMacRSIItemIndex(ByVal code As String) As Long
    Dim result As Long
    If Len(code) <> 2 Then Exit Function
    If InStr(1, "0123456789", Left$(code, 1), vbBinaryCompare) = 0 Then Exit Function
    If InStr(1, "0123456789", Right$(code, 1), vbBinaryCompare) = 0 Then Exit Function
    result = CLng(code)
    If result >= 1 And result <= 12 Then smfMacRSIItemIndex = result
End Function

Private Function smfMacRSITicker(ByVal ticker As String) As String
    Dim i As Long, ch As String, result As String
    If Len(ticker) < 1 Or Len(ticker) > 64 Then Exit Function
    For i = 1 To Len(ticker)
        ch = Mid$(ticker, i, 1)
        Select Case ch
            Case "A" To "Z", "0" To "9", "-", ".", "_": result = result & ch
            Case "^": result = result & "%5E"
            Case "=": result = result & "%3D"
            Case Else: Exit Function
        End Select
    Next i
    smfMacRSITicker = result
End Function

Private Function smfMacRSINumber(ByVal value As Variant, ByRef number As Double) As Boolean
    On Error GoTo Invalid
    Select Case VarType(value)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            number = CDbl(value)
            If Abs(number) > 1E+100 Then Exit Function
            smfMacRSINumber = True
    End Select
Invalid:
End Function

Private Function smfMacRSIDate(ByVal stamp As Double, ByVal offset As Double) As Date
    On Error GoTo Invalid
    If stamp <= 0 Or stamp > 253402214400# Then Exit Function
    smfMacRSIDate = DateSerial(1970, 1, 1) + Int((stamp + offset) / 86400#)
Invalid:
End Function

Private Function smfMacRSIRatio(ByVal gain As Double, ByVal loss As Double) As Double
    If gain + loss = 0 Then
        smfMacRSIRatio = 50
    Else
        smfMacRSIRatio = 100 * gain / (gain + loss)
    End If
End Function

Private Function smfMacRSIHistory(ByVal snapshot As Object, ByVal quoteDate As Date, _
        ByVal offset As Double, ByVal requiredCloses As Long, ByRef previousClose As Double, _
        ByRef averageGain As Double, ByRef averageLoss As Double, ByRef bars As Object, _
        ByRef currentBar As Long) As Boolean
    Dim stamps As Object, closes As Object, values(1 To 21) As Double
    Dim i As Long, count As Long, stamp As Double, closeValue As Double
    Dim day As Date, laterDay As Date, change As Double
    On Error GoTo Invalid
    Set stamps = snapshot("timestamp")
    Set bars = snapshot("indicators")("quote")(1)
    If stamps.Count < 1 Or stamps.Count > 500 Then Exit Function
    If requiredCloses > 0 Then
        Set closes = bars("close")
        If closes.Count <> stamps.Count Then Exit Function
    End If
    laterDay = DateSerial(9999, 12, 31)
    For i = stamps.Count To 1 Step -1
        If Not smfMacRSINumber(stamps(i), stamp) Then Exit Function
        day = smfMacRSIDate(stamp, offset)
        If day = 0 Or day >= laterDay Then Exit Function
        laterDay = day
        If day = quoteDate Then currentBar = i
        If day > quoteDate Then
            ' A numeric future bar and an older quote are different snapshots.
            If requiredCloses > 0 Then
                If smfMacRSINumber(closes(i), closeValue) Then
                    If closeValue > 0 Then Exit Function
                End If
            End If
        ElseIf day < quoteDate Then
            If requiredCloses > 0 Then
                If Not smfMacRSINumber(closes(i), closeValue) Then Exit Function
                If closeValue <= 0 Then Exit Function
                count = count + 1
                values(count) = closeValue
                If count = requiredCloses Then Exit For
            End If
        End If
    Next i
    If count < requiredCloses Then Exit Function
    If requiredCloses > 0 Then previousClose = values(1)
    If requiredCloses = 21 Then
        For i = 20 To 19 Step -1
            change = values(i) - values(i + 1)
            If change > 0 Then averageGain = averageGain + change / 2
            If change < 0 Then averageLoss = averageLoss - change / 2
        Next i
        For i = 18 To 1 Step -1
            change = values(i) - values(i + 1)
            averageGain = averageGain / 2
            averageLoss = averageLoss / 2
            If change > 0 Then averageGain = averageGain + change / 2
            If change < 0 Then averageLoss = averageLoss - change / 2
        Next i
    End If
    smfMacRSIHistory = True
Invalid:
End Function

Private Function smfMacRSIOptionalNumber(ByVal data As Object, ByVal field As String, _
        ByVal positive As Boolean) As Variant
    Dim number As Double
    smfMacRSIOptionalNumber = "--"
    On Error GoTo Missing
    If Not data.Exists(field) Then Exit Function
    If Not smfMacRSINumber(data(field), number) Then Exit Function
    If number < 0 Then Exit Function
    If positive Then
        If number = 0 Then Exit Function
    End If
    smfMacRSIOptionalNumber = number
Missing:
End Function

Private Function smfMacRSIBarNumber(ByVal bars As Object, ByVal field As String, _
        ByVal row As Long, ByVal positive As Boolean) As Variant
    Dim data As Object, number As Double
    smfMacRSIBarNumber = "--"
    On Error GoTo Missing
    If row < 1 Then Exit Function
    Set data = bars(field)
    If row > data.Count Then Exit Function
    If Not smfMacRSINumber(data(row), number) Then Exit Function
    If number < 0 Then Exit Function
    If positive Then
        If number = 0 Then Exit Function
    End If
    smfMacRSIBarNumber = number
Missing:
End Function

Private Sub smfMacRSIBidAsk(ByVal ticker As String, ByVal encoded As String, _
        ByVal quoteDate As Date, ByVal offset As Double, ByRef bid As Variant, ByRef ask As Variant)
    Dim payload As String, root As Object, quote As Object, stamp As Double
    On Error GoTo Missing
    payload = smfGetWebPage("https://query1.finance.yahoo.com/v7/finance/quote?fields=symbol,regularMarketTime,bid,ask&formatted=false&symbols=" & encoded, 4, 0)
    Set root = smfMacRSIParseJSON(payload)
    If Not IsNull(root("quoteResponse")("error")) Then Exit Sub
    If root("quoteResponse")("result").Count <> 1 Then Exit Sub
    Set quote = root("quoteResponse")("result")(1)
    If UCase$(CStr(quote("symbol"))) <> ticker Then Exit Sub
    If Not smfMacRSINumber(quote("regularMarketTime"), stamp) Then Exit Sub
    If smfMacRSIDate(stamp, offset) <> quoteDate Then Exit Sub
    bid = smfMacRSIOptionalNumber(quote, "bid", False)
    ask = smfMacRSIOptionalNumber(quote, "ask", False)
Missing:
End Sub
#End If

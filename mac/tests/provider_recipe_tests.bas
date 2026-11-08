Attribute VB_Name = "modProviderRecipeTests"
Option Explicit

#If Mac Then
' Import only into an isolated diagnostic workbook, not the production add-in.
' Copy the updated element files beside that workbook and load them first.
' Public page fragments contain data and markup only. No cookies or scripts.
Public Function SMFMacProviderRecipeSelfTest() As Long
    Dim elements(1 To 8) As Integer, kinds(1 To 8) As Integer
    Dim tickers(1 To 8) As String, urls(1 To 8) As String, html(1 To 8) As String
    Dim expected(1 To 8) As Variant, actual As Variant, recipe As String, initialized As Variant
    Dim oldKeys(1 To kPages) As String, oldBodies(1 To kPages) As String
    Dim oldTexts(1 To 3) As String, oldCache As String, oldLog As String, oldError As Variant
    Dim beforeURL As String, beforeError As String, beforeStatus As Long
    Dim i As Long, checks As Long, accepted As Boolean, savedState As Boolean
    Dim errorNumber As Long, errorText As String
    On Error GoTo TestFailed
    initialized = RCHGetElementNumber("Definition", 4074)
    If CStr(RCHGetElementNumber("Definition", 5196)) <> "AdvFN-A;Annual Indicators -- Ending Date -- FY1;=DATEVALUE(smfStrExtr(smfGetTagContent(""https://www.advfn.com/p.php?pid=financials&symbol=~~~~~"",""TD"",-1,"">Year End""),""Year End "",""~""))" Then _
        Err.Raise 5, , "Load the updated local element definition for 5196 before this test"
    If CStr(RCHGetElementNumber("Definition", 849)) <> "Zacks;Target Price Consensus;=RCHGetTableCell(""https://www.zacks.com/stock/research/~~~~~/price-target-stock-forecast"",1,"">Average Price Target"",,,,1)" Then _
        Err.Raise 5, , "Load the updated local element definition for 849 before this test"
    If CStr(RCHGetElementNumber("Definition", 19004)) <> "NetFlix;Year;=smfGetTagContent(""https://www.netflix.com/title/~~~~~"",""li"",1,""data-uia=""""metadata"""""",,,,1)" Then _
        Err.Raise 5, , "Load the updated local element definition for 19004 before this test"
    If CStr(RCHGetElementNumber("Definition", 2727)) <> "GuruDiv;Dividend Yield;=smfGetTagContent(""https://www.gurufocus.com/stock/~~~~~/dividend"",""div"",-1,"" (Forward)"",,,,1,,0,32767,"" (Forward)"")" Then _
        Err.Raise 5, , "Load the updated local element definition for 2727 before this test"
    beforeURL = SMFMacLastURL(): beforeError = SMFMacLastError(): beforeStatus = SMFMacHTTPStatus()
    oldCache = sWebCache: oldLog = sLog: oldError = vError
    For i = 1 To kPages
        oldKeys(i) = aData(i, 1): oldBodies(i) = aData(i, 2)
    Next i
    For i = 1 To 3: oldTexts(i) = sData(i): Next i
    savedState = True
    Erase aData
    sWebCache = "Y": sLog = "N"
    elements(1) = 2727: kinds(1) = 1
    tickers(1) = "AAPL": urls(1) = "https://www.gurufocus.com/stock/AAPL/dividend"
    expected(1) = 0.0034#
    html(1) = html(1) & "<html><body><div><div class=""el-row"" data-v-8b86a846>0.34% (Trailing)</div><div class=""m-v-xs el-row"" data-v-8b86a846>0.34% (Forward)</div><h4>Dividen"
    html(1) = html(1) & "d Yield</h4></div></body></html>"
    elements(2) = 2727: kinds(2) = 1
    tickers(2) = "MSFT": urls(2) = "https://www.gurufocus.com/stock/MSFT/dividend"
    expected(2) = 0.0078#
    html(2) = html(2) & "<html><body><div><div class=""el-row"" data-v-8b86a846>0.71% (Trailing)</div><div class=""m-v-xs el-row"" data-v-8b86a846>0.78% (Forward)</div><h4>Dividen"
    html(2) = html(2) & "d Yield</h4></div></body></html>"
    elements(3) = 849: kinds(3) = 2
    tickers(3) = "AAPL": urls(3) = "https://www.zacks.com/stock/research/AAPL/price-target-stock-forecast"
    expected(3) = 338.68#
    html(3) = html(3) & "<html><body><table> <thead> <tr> <th scope=""col"" class=""align_center"">Average Price Target</th> <th scope=""col"" class=""align_center"">Highest Price Tar"
    html(3) = html(3) & "get</th> <th scope=""col"" class=""align_center"">Lowest Price Target</th> <th scope=""col"" class=""align_center"">Downside to Average Price Target</th> </tr"
    html(3) = html(3) & "> </thead> <tbody> <tr> <th scope=""col"" class=""align_center"">$338.68</th> <td class=""align_center"">$400.00</td> <td class=""align_center"">$250.00</th> "
    html(3) = html(3) & "<td class=""align_center"">-1.60%</td> </tr> </tbody> </table></body></html>"
    elements(4) = 849: kinds(4) = 2
    tickers(4) = "MSFT": urls(4) = "https://www.zacks.com/stock/research/MSFT/price-target-stock-forecast"
    expected(4) = 629.46#
    html(4) = html(4) & "<html><body><table> <thead> <tr> <th scope=""col"" class=""align_center"">Average Price Target</th> <th scope=""col"" class=""align_center"">Highest Price Tar"
    html(4) = html(4) & "get</th> <th scope=""col"" class=""align_center"">Lowest Price Target</th> <th scope=""col"" class=""align_center"">Upside to Average Price Target</th> </tr> "
    html(4) = html(4) & "</thead> <tbody> <tr> <th scope=""col"" class=""align_center"">$629.46</th> <td class=""align_center"">$750.00</td> <td class=""align_center"">$440.00</th> <t"
    html(4) = html(4) & "d class=""align_center"">17.45%</td> </tr> </tbody> </table></body></html>"
    elements(5) = 19004: kinds(5) = 1
    tickers(5) = "70111470": urls(5) = "https://www.netflix.com/title/70111470"
    expected(5) = 2010#
    html(5) = html(5) & "<html><body><div data-uia=""metadata""><ul><li class=""default-ltr-iqcdef-cache-6prs41 epf12f60"">2010</li></ul></div></body></html>"
    elements(6) = 19004: kinds(6) = 1
    tickers(6) = "80175798": urls(6) = "https://www.netflix.com/title/80175798"
    expected(6) = 2019#
    html(6) = html(6) & "<html><body><div data-uia=""metadata""><ul><li class=""default-ltr-iqcdef-cache-6prs41 epf12f60"">2019</li></ul></div></body></html>"
    elements(7) = 5196: kinds(7) = 3
    tickers(7) = "NASDAQ:AAPL": urls(7) = "https://www.advfn.com/p.php?pid=financials&symbol=NASDAQ:AAPL"
    expected(7) = DateSerial(2025, 9, 27)
    html(7) = html(7) & "<html><body><table><tr> <th class=""number-field-small sticky-col first-col small"">Year End September 27 2025</th> <th class=""number-field"">2021</th><t"
    html(7) = html(7) & "h class=""number-field"">2022</th><th class=""number-field"">2023</th><th class=""number-field"">2024</th><th class=""number-field"">2025</th> </tr></table></"
    html(7) = html(7) & "body></html>"
    elements(8) = 5196: kinds(8) = 3
    tickers(8) = "NASDAQ:MSFT": urls(8) = "https://www.advfn.com/p.php?pid=financials&symbol=NASDAQ:MSFT"
    expected(8) = DateSerial(2026, 6, 30)
    html(8) = html(8) & "<html><body><table><tr> <th class=""number-field-small sticky-col first-col small"">Year End June 30 2026</th> <th class=""number-field"">2022</th><th cla"
    html(8) = html(8) & "ss=""number-field"">2023</th><th class=""number-field"">2024</th><th class=""number-field"">2025</th><th class=""number-field"">2026</th> </tr></table></body>"
    html(8) = html(8) & "</html>"
    For i = 1 To 8
        aData(i, 1) = "0:" & urls(i)
        aData(i, 2) = ProviderFixtureNormalize(html(i))
    Next i
    For i = 1 To 8
        recipe = Replace(CStr(RCHGetElementNumber("P-URL", elements(i))), "~~~~~", tickers(i))
        actual = "unchanged": accepted = False
        Select Case kinds(i)
            Case 1: accepted = smfMacTryTagContentRecipe(recipe, actual)
            Case 2: accepted = smfMacTryTableCellRecipe(recipe, actual)
            Case 3: accepted = smfMacTryNestedRecipe(recipe, actual)
        End Select
        If Not accepted Then Err.Raise 5, , "Dispatcher rejected fixture " & CStr(elements(i)) & " " & tickers(i)
        ProviderFixtureAssert actual, expected(i), CStr(elements(i)) & " direct " & tickers(i)
        checks = checks + 1
        actual = RCHGetElementNumber(tickers(i), elements(i), "Error")
        ProviderFixtureAssert actual, expected(i), CStr(elements(i)) & " numbered " & tickers(i)
        checks = checks + 1
    Next i
    For i = 1 To 8
        aData(i, 2) = "<html><body>No requested provider field</body></html>"
    Next i
    For i = 1 To 8
        actual = RCHGetElementNumber(tickers(i), elements(i), "Error")
        If VarType(actual) <> vbString Then Err.Raise 5, , "Missing fixture field became a numeric value"
        If CStr(actual) <> "Error" Then Err.Raise 5, , "Missing fixture field lost Error for " & CStr(elements(i))
        checks = checks + 1
    Next i
    If SMFMacLastURL() <> beforeURL Or SMFMacLastError() <> beforeError Or _
       SMFMacHTTPStatus() <> beforeStatus Then Err.Raise 5, , "Cached provider fixtures made an HTTP request"
    checks = checks + 1
    If checks <> 25 Then Err.Raise 5, , "Unexpected provider fixture count"
    SMFMacProviderRecipeSelfTest = checks
    GoTo RestoreState
TestFailed:
    errorNumber = Err.Number: errorText = Err.Description
    Resume RestoreState
RestoreState:
    On Error Resume Next
    If savedState Then
        For i = 1 To kPages
            aData(i, 1) = oldKeys(i): aData(i, 2) = oldBodies(i)
        Next i
        For i = 1 To 3: sData(i) = oldTexts(i): Next i
        sWebCache = oldCache: sLog = oldLog: vError = oldError
    End If
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, "SMFMacProviderRecipeSelfTest", errorText
End Function

Private Function ProviderFixtureNormalize(ByVal html As String) As String
    ' Match the table-header conversion at the normal HTTP-to-cache boundary.
    html = Replace(html, "<th", "<td"): html = Replace(html, "</th", "</td")
    html = Replace(html, "<TH", "<td"): html = Replace(html, "</TH", "</td")
    ProviderFixtureNormalize = html
End Function

Private Sub ProviderFixtureAssert(ByVal actual As Variant, ByVal expected As Variant, ByVal label As String)
    If IsError(actual) Then Err.Raise 5, , "Excel error for " & label
    If VarType(actual) <> VarType(expected) Then _
        Err.Raise 5, , "Wrong value type for " & label & ": " & TypeName(actual)
    If Abs(CDbl(actual) - CDbl(expected)) > 0.000000001 Then _
        Err.Raise 5, , "Wrong value for " & label & ": " & CStr(actual)
End Sub
#End If

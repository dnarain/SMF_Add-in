Attribute VB_Name = "smfCreateComment_"
Public Function RCHCreateComment(pTicker As String, _
                  Optional ByVal pChoice As Integer = 1, _
                  Optional ByVal pWidth As Integer = 0, _
                  Optional ByVal pHeight As Integer = 0, _
                  Optional ByVal pVisible As Integer = 0, _
                  Optional ByVal pTop As Integer = 1, _
                  Optional ByVal pLeft As Integer = 1, _
                  Optional ByVal pScale As Single = 1#, _
                  Optional ByVal pText As String = "", _
                  Optional ByVal pReturn As String = "Chart")
Attribute RCHCreateComment.VB_Description = "Creates a comment box for the cell that can contain text and/or an image (e.g. a stock chart). "

    '-----------------------------------------------------------------------------------------------------------*
    ' Function to create a comment object and insert image/text into it
    '-----------------------------------------------------------------------------------------------------------*
    ' 2007.01.17 -- Change CCur() usage to CDec() because of precision issues
    ' 2007.09.24 -- Set the comment line color to white in order to "hide" it
    ' 2009.06.26 -- Add ActiveWorkbook.Name error check
    ' 2016.05.18 -- Change CDec() to smfCDec() to ease transition between operating systems
    ' 2018.01.24 -- Change AdvFN URL from "http://" to "https://"
    ' 2018.12.27 -- Change "http://stockcharts.com" domains to "https://c.stockcharts.com"
    '-----------------------------------------------------------------------------------------------------------*
    ' Examples of usage:
    '
    '    =RCHCreateComment("MMM")
    '    =RCHCreateComment("MMM",1,350,390,1,1,1)
    '
    '-----------------------------------------------------------------------------------------------------------*
    
#If Mac Then
    Dim macCaller As Object
    On Error Resume Next
    Set macCaller = Application.Caller
    On Error GoTo 0
    RCHCreateComment = smfMacCreateComment(macCaller, pTicker, pChoice, pWidth, pHeight, _
                                         pVisible, pTop, pLeft, pScale, pText, pReturn)
    Exit Function
#Else
    Set oCell = Cells(Application.Caller.Cells.Row, Application.Caller.Cells.Column)
    If ActiveWorkbook.Name <> Application.Caller.Parent.Parent.Name Then Exit Function
    If ActiveSheet.Name <> Application.Caller.Worksheet.Name Then Exit Function
    On Error Resume Next
    oCell.Comment.Delete
    On Error GoTo 0
    Select Case True
       Case UCase(pTicker) = "NONE": RCHCreateComment = "None": Exit Function
       Case pChoice = 0
            sURL = ""
            If pWidth = 0 Then pWidth = 300
            If pHeight = 0 Then pHeight = 200
       Case pTicker = "": GoTo ErrorExit
       Case pScale <= 0: GoTo ErrorExit
       Case pWidth < 0: GoTo ErrorExit
       Case pHeight < 0: GoTo ErrorExit
       Case pChoice = 1      ' Daily Chart of Gallery View from StockCharts
            sURL = "https://c.stockcharts.com/c-sc/sc?chart=" & pTicker & ",uu[h,a]daclyyay[pb50!b200!f][vc60][iue12,26,9!lc20]"
            If pWidth = 0 Then pWidth = 350 * pScale
            If pHeight = 0 Then pHeight = 390 * pScale
       Case pChoice = 2      ' P&F Chart from StockCharts
            sURL = "https://stockcharts.com/def/servlet/SharpChartv05.ServletDriver?chart=" & pTicker & ",pltad[pa][da][f!3!!]&pnf=y"
            If pWidth = 0 Then pWidth = 390 * pScale
            If pHeight = 0 Then pHeight = 314 * pScale
       Case pChoice = 3      ' 6-month Candleglance Chart from StockCharts
            sURL = "https://c.stockcharts.com/c-sc/sc?chart=" & pTicker & ",uu[305,a]dacayaci[pb20!b50][dc]"
            If pWidth = 0 Then pWidth = 229 * pScale
            If pHeight = 0 Then pHeight = 132 * pScale
       Case pChoice = 4      ' 6-month chart from Business Week Online
            sURL = "https://c.stockcharts.com/c-sc/sc?s=" & pTicker & "&p=D&yr=0&mn=6&dy=0&i=t94339682869&r=4806"
            If pWidth = 0 Then pWidth = 638 * pScale
            If pHeight = 0 Then pHeight = 501 * pScale
       Case pChoice = 5      ' 6-month chart Rule #1 Technicals from StockCharts
            sURL = "https://c.stockcharts.com/c-sc/sc?s=" & pTicker & "&p=D&yr=0&mn=6&dy=0&i=t39628903145&r=9933"
            If pWidth = 0 Then pWidth = 350 * pScale
            If pHeight = 0 Then pHeight = 360 * pScale
       Case pChoice = 97
            sURL = "https://www.advfn.com/p.php?pid=financialgraphs" '&a0=13&a1=13&a2=10&a3=8&a4=8&a5=10"
            aSplit = Split(pTicker, ",")
            iNbr = UBound(aSplit, 1)
            If iNbr <> 4 Then GoTo ErrorExit
            For i1 = 0 To iNbr
                sURL = sURL & "&a" & i1 & "=" & aSplit(i1)
                Next i1
            If pWidth = 0 Then pWidth = 263 * pScale
            If pHeight = 0 Then pHeight = 169 * pScale
       Case pChoice = 98
            sURL = "http://ogres-crypt.com/php/chart.php?d="
            aSplit = Split(pTicker, ",")
            iNbr = UBound(aSplit, 1)
            iMax = 0.01
            iMin = 999999999
            For i1 = 0 To iNbr
                iTemp = 0
                On Error Resume Next
                iTemp = smfCDec(aSplit(i1))
                On Error GoTo 0
                If (iTemp > iMax) Then iMax = iTemp
                If (iTemp < iMin And iTemp > 0) Then iMin = iTemp
                Next i1
            For i1 = 0 To iNbr
                iTemp = 0
                On Error Resume Next
                iTemp = smfCDec(aSplit(i1))
                On Error GoTo 0
                iTemp = IIf(iTemp > 0, 1 + 97 * (iTemp - iMin) / (iMax - iMin), 0)
                sURL = sURL & CInt(iTemp) & IIf(i1 = iNbr, "", ",")
                Next i1
            If pWidth = 0 Then pWidth = 36 * iNbr * pScale
            If pHeight = 0 Then pHeight = 90 * pScale
       Case pChoice = 99
            sURL = pTicker
            If pWidth = 0 Then pWidth = 400
            If pHeight = 0 Then pHeight = 300
       Case Else: GoTo ErrorExit
       End Select
    oCell.AddComment ("")
    If sURL <> "" Then oCell.Comment.Shape.Fill.UserPicture sURL
    oCell.Comment.Text Text:=IIf(pText = "", Chr(32), pText)
    oCell.Comment.Shape.Width = pWidth
    oCell.Comment.Shape.Height = pHeight
    oCell.Comment.Shape.Top = pTop + oCell.Top
    oCell.Comment.Shape.Left = pLeft + oCell.Left
    oCell.Comment.Shape.Line.Visible = False             ' Doesn't work
    oCell.Comment.Shape.Line.ForeColor.SchemeColor = 9   ' Set line color to background color
    oCell.Comment.Shape.Shadow.Visible = False
    oCell.Comment.Visible = IIf(pVisible = 1, True, False)
    RCHCreateComment = pReturn
    Exit Function
ErrorExit:
    RCHCreateComment = "Error"
#End If
    End Function



#If Mac Then
' Excel for Mac 16.114.4 raised445 for Comment.Shape.Fill. Keep note text
' and embed a separate worksheet picture, keyed to its formula cell.
Private Function smfMacCreateComment(ByVal caller As Variant, ByVal ticker As String, _
        ByVal choice As Integer, ByVal width As Integer, ByVal height As Integer, _
        ByVal visible As Integer, ByVal offsetTop As Integer, ByVal offsetLeft As Integer, _
        ByVal pictureScale As Single, ByVal noteText As String, ByVal returnText As String) As Variant
    Dim target As Range, picture As Shape, numericChart As ChartObject
    Dim source As String, imageFile As String, ownedDownload As Boolean
    Dim failure As String, reply As String, parts As Variant, operation As String
    operation = "Preparing the chart cell"
    On Error GoTo Failed
    If Not IsObject(caller) Then GoTo InvalidInput
    If TypeName(caller) <> "Range" Then GoTo InvalidInput
    Set target = caller.Cells(1, 1)
    If ActiveWorkbook.Name <> target.Parent.Parent.Name Then Exit Function
    If ActiveSheet.Name <> target.Worksheet.Name Then Exit Function
    operation = "Removing the previous chart"
    Call smfMacRemoveCommentPicture(target)
    On Error Resume Next
    target.Comment.Delete
    On Error GoTo Failed
    If UCase$(ticker) = "NONE" Then
        smfMacCreateComment = "None"
        Exit Function
    End If
    If pictureScale <= 0 Or width < 0 Or height < 0 Then GoTo InvalidInput
    source = smfMacCommentImageURL(ticker, choice, width, height, pictureScale)
    If choice <> 0 And source = "" Then GoTo InvalidInput
    operation = "Creating the cell note"
    If choice = 0 Then
        target.AddComment IIf(noteText = "", " ", noteText)
    Else
        target.AddComment IIf(noteText = "", "", noteText & vbLf) & _
            "SMF stores this chart on the worksheet. Select this cell and use " & _
            "SMF: Show/Hide Cell Chart in the cell menu."
    End If
    ' Cosmetic comment properties vary between Excel for Mac builds. A missing
    ' border or shadow operation must not turn a valid text note into #VALUE!.
    On Error Resume Next
    target.Comment.Shape.Width = width
    target.Comment.Shape.Height = height
    target.Comment.Shape.Top = target.Top + offsetTop
    target.Comment.Shape.Left = target.Left + offsetLeft
    target.Comment.Visible = (visible = 1)
    On Error GoTo Failed
    If choice = 0 Then
        smfMacCreateComment = returnText
        Exit Function
    End If
    operation = "Creating the local numeric chart"
    If choice = 98 Then
        Set numericChart = target.Worksheet.ChartObjects.Add(target.Left + offsetLeft, _
                                target.Top + offsetTop, width, height)
        numericChart.Chart.ChartType = xlLine
        numericChart.Chart.SeriesCollection.NewSeries.Values = smfMacNumericChartValues(ticker)
        numericChart.Chart.HasLegend = False
        On Error Resume Next
        numericChart.Chart.Axes(xlValue).MinimumScale = 0
        numericChart.Chart.Axes(xlValue).MaximumScale = 100
        On Error GoTo Failed
        Set picture = target.Worksheet.Shapes(numericChart.Name)
    Else
    If LCase$(Left$(source, 7)) = "http://" Or LCase$(Left$(source, 8)) = "https://" Then
        operation = "Downloading the image"
        reply = AppleScriptTask("SMFMac.scpt", "downloadImage", source)
        parts = Split(reply, vbLf)
        If UBound(parts) < 3 Then Err.Raise 5, "RCHCreateComment", "The image helper returned an invalid response"
        If parts(0) <> "SMFIMAGE/1" Then Err.Raise 5, "RCHCreateComment", "The image helper returned an invalid response"
        If parts(1) <> "200" Or parts(2) = "" Then
            failure = CStr(parts(3))
            If failure = "" Then failure = "The image request did not return a picture"
            GoTo FailedImage
        End If
        imageFile = CStr(parts(2))
        ownedDownload = True
    ElseIf Left$(source, 1) = "/" Then
        operation = "Preparing the local image"
        reply = AppleScriptTask("SMFMac.scpt", "prepareLocalImage", source)
        parts = Split(reply, vbLf)
        If UBound(parts) < 3 Then Err.Raise 5, "RCHCreateComment", "The image helper returned an invalid response"
        If parts(0) <> "SMFIMAGE/1" Then Err.Raise 5, "RCHCreateComment", "The image helper returned an invalid response"
        If parts(1) <> "200" Or parts(2) = "" Then
            failure = CStr(parts(3))
            If failure = "" Then failure = "The local file could not be prepared as a picture"
            GoTo FailedImage
        End If
        imageFile = CStr(parts(2))
        ownedDownload = True
    Else
        Err.Raise 5, "RCHCreateComment", "Use an HTTP image URL or an absolute Mac file path"
    End If
    operation = "Embedding the image"
    Set picture = target.Worksheet.Shapes.AddPicture(imageFile, msoFalse, msoTrue, _
                    target.Left + offsetLeft, target.Top + offsetTop, width, height)
    End If
    operation = "Naming the chart"
    picture.Name = smfMacCommentPictureName(target)
    operation = "Setting chart metadata"
    picture.AlternativeText = "SMFCommentPicture/1" & vbLf & target.Address(False, False) & vbLf & source
    operation = "Anchoring the chart to its cell"
    picture.Placement = xlMoveAndSize
    operation = "Setting chart visibility"
    picture.Visible = IIf(visible = 1, msoTrue, msoFalse)
    If ownedDownload Then reply = AppleScriptTask("SMFMac.scpt", "releaseImage", imageFile)
    smfMacCreateComment = returnText
    Exit Function
InvalidInput:
    operation = "Checking chart arguments"
    failure = "The chart arguments are invalid"
    GoTo FailedImage
Failed:
    failure = CStr(Err.Number) & ": " & Err.Description
FailedImage:
    If operation <> "" Then failure = operation & ": " & failure
    On Error Resume Next
    If Not picture Is Nothing Then picture.Delete
    If Not numericChart Is Nothing Then numericChart.Delete
    If ownedDownload Then reply = AppleScriptTask("SMFMac.scpt", "releaseImage", imageFile)
    If Not target Is Nothing Then
        If target.Comment Is Nothing Then target.AddComment " "
        target.Comment.Text IIf(noteText = "", "", noteText & vbLf) & "SMF chart unavailable: " & failure
    End If
    smfMacCreateComment = "Error"
End Function

Private Function smfMacCommentPictureName(ByVal target As Range) As String
    smfMacCommentPictureName = "SMFCommentPicture_" & CStr(target.Row) & "_" & CStr(target.Column)
End Function

Private Sub smfMacRemoveCommentPicture(ByVal target As Range)
    Dim picture As Shape
    On Error Resume Next
    Set picture = target.Worksheet.Shapes(smfMacCommentPictureName(target))
    On Error GoTo 0
    If picture Is Nothing Then Exit Sub
    If Left$(picture.AlternativeText, Len("SMFCommentPicture/1" & vbLf)) <> "SMFCommentPicture/1" & vbLf Then
        Err.Raise 5, "RCHCreateComment", "Another worksheet object uses this chart name"
    End If
    picture.Delete
End Sub

Public Sub smfToggleCommentPicture()
    Dim target As Range, picture As Shape
    On Error GoTo Failed
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set target = Selection.Cells(1, 1)
    Set picture = target.Worksheet.Shapes(smfMacCommentPictureName(target))
    If Left$(picture.AlternativeText, Len("SMFCommentPicture/1" & vbLf)) <> "SMFCommentPicture/1" & vbLf Then Exit Sub
    picture.Visible = IIf(picture.Visible = msoFalse, msoTrue, msoFalse)
    Exit Sub
Failed:
    MsgBox "This cell has no SMF chart to show. Check the formula result and the cell note.", vbInformation, "SMF chart"
End Sub

Private Function smfMacCommentImageURL(ByVal ticker As String, ByVal choice As Integer, _
        ByRef width As Integer, ByRef height As Integer, ByVal pictureScale As Single) As String
    Dim source As String, values As Variant, i As Long, count As Long
    Select Case choice
        Case 0
            If width = 0 Then width = 300
            If height = 0 Then height = 200
        Case 1
            source = "https://stockcharts.com/c-sc/sc?chart=" & ticker & ",uu[h,a]daclyyay[pb50!b200!f][vc60][iue12,26,9!lc20]"
            If width = 0 Then width = 350 * pictureScale
            If height = 0 Then height = 390 * pictureScale
        Case 2
            source = "https://stockcharts.com/pnf/chart?c=" & ticker & ",P"
            If width = 0 Then width = 390 * pictureScale
            If height = 0 Then height = 314 * pictureScale
        Case 3
            source = "https://stockcharts.com/c-sc/sc?chart=" & ticker & ",uu[305,a]dacayaci[pb20!b50][dc]"
            If width = 0 Then width = 229 * pictureScale
            If height = 0 Then height = 132 * pictureScale
        Case 4
            source = "https://stockcharts.com/c-sc/sc?s=" & ticker & "&p=D&yr=0&mn=6&dy=0&i=t94339682869&r=4806"
            If width = 0 Then width = 638 * pictureScale
            If height = 0 Then height = 501 * pictureScale
        Case 5
            source = "https://stockcharts.com/c-sc/sc?s=" & ticker & "&p=D&yr=0&mn=6&dy=0&i=t39628903145&r=9933"
            If width = 0 Then width = 350 * pictureScale
            If height = 0 Then height = 360 * pictureScale
        Case 97
            values = Split(ticker, ",")
            If UBound(values) <> 4 Then Exit Function
            source = "https://www.advfn.com/p.php?pid=financialgraphs"
            For i = 0 To UBound(values)
                source = source & "&a" & CStr(i) & "=" & CStr(values(i))
            Next i
            If width = 0 Then width = 263 * pictureScale
            If height = 0 Then height = 169 * pictureScale
        Case 98
            values = Split(ticker, ",")
            count = UBound(values)
            source = "SMF local numeric chart: " & ticker
            If width = 0 Then width = 36 * IIf(count < 1, 1, count) * pictureScale
            If height = 0 Then height = 90 * pictureScale
        Case 99
            source = ticker
            If width = 0 Then width = 400
            If height = 0 Then height = 300
    End Select
    smfMacCommentImageURL = source
End Function
Private Function smfMacNumericChartValues(ByVal ticker As String) As Variant
    Dim values As Variant, plotted() As Double
    Dim maximum As Double, minimum As Double, value As Double, i As Long
    Dim hasPositive As Boolean
    values = Split(ticker, ",")
    ReDim plotted(0 To UBound(values))
    For i = 0 To UBound(values)
        value = 0
        On Error Resume Next
        value = smfCDec(values(i))
        On Error GoTo 0
        If value > 0 Then
            If Not hasPositive Then
                minimum = value
                maximum = value
                hasPositive = True
            Else
                If value > maximum Then maximum = value
                If value < minimum Then minimum = value
            End If
        End If
    Next i
    For i = 0 To UBound(values)
        value = 0
        On Error Resume Next
        value = smfCDec(values(i))
        On Error GoTo 0
        If value > 0 Then
            If maximum = minimum Then
                plotted(i) = 50
            Else
                plotted(i) = CInt(1 + 97 * (value - minimum) / (maximum - minimum))
            End If
        Else
            plotted(i) = 0
        End If
    Next i
    smfMacNumericChartValues = plotted
End Function
#End If

Attribute VB_Name = "smfGetGoogleHistoryCSV_"
Option Explicit

Function smfGetGoogleHistoryCSV(ByVal pTicker As String, _
                   Optional ByVal pStartDate As Variant = "", _
                   Optional ByVal pEndDate As Variant = "", _
                   Optional ByVal pPeriod As String = "d", _
                   Optional ByVal pRows As Integer = 10000, _
                   Optional ByVal pCols As Integer = 7)

    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to process CSV file from Google
    '-----------------------------------------------------------------------------------------------------------*
    ' 2017.06.12 -- Added
    ' 2017.11.30 -- Change URL ("www" to "finance"
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
    Dim dBegin As Variant, dEnd As Variant
    vData(1, 1) = "Error on starting date: " & pStartDate
    Select Case True
          Case VarType(pStartDate) = vbDate Or VarType(pStartDate) = vbDouble
               dBegin = pStartDate
          Case pStartDate = ""
               dBegin = DateValue("1/1/1970")
          Case Else
               dBegin = DateValue(pStartDate)
          End Select
    vData(1, 1) = "Error on ending date: " & pEndDate
    Select Case True
          Case VarType(pEndDate) = vbDate Or VarType(pEndDate) = vbDouble
               dEnd = Int(pEndDate)
          Case pEndDate = ""
               dEnd = Int(Now)
          Case Else
               dEnd = Int(DateValue(pEndDate))
          End Select
     If dBegin > dEnd Then
        vData(1, 1) = "Error: Starting date cannot be after ending date: " & pStartDate & "," & pEndDate
        GoTo ErrorExit
        End If
       
    '------------------> Process period
    ' No processing at this point
    
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
#If Mac Then
    ' Encode each query value so dates, tickers and separators form one valid URL.
    sURL = "http://finance.google.com/finance/historical?output=csv&q=" & smfMacGoogleHistoryQuery(UCase(pTicker)) & _
           "&startdate=" & smfMacGoogleHistoryQuery(Format(dBegin, "mmm d, yyyy")) & _
           "&enddate=" & smfMacGoogleHistoryQuery(Format(dEnd, "mmm d, yyyy"))
#Else
    sURL = "http://finance.google.com/finance/historical?output=csv&q=" & UCase(pTicker) & _
           "&startdate=" & Format(dBegin, "mmm d, yyyy") & "&enddate=" & Format(dEnd, "mmm d, yyyy")
#End If
    vData = smfGetCSVFile(sURL, ",", iRows, iCols)

ErrorExit:
    smfGetGoogleHistoryCSV = vData
                   
   End Function


#If Mac Then
' Percent-encode UTF-8 query values without changing their characters or order.
Private Function smfMacGoogleHistoryQuery(ByVal pValue As String) As String
    Dim i As Long, code As Long, low As Long, result As String
    For i = 1 To Len(pValue)
        code = AscW(Mid$(pValue, i, 1)) And &HFFFF&
        If code >= &HD800& And code <= &HDBFF& Then
            If i = Len(pValue) Then Err.Raise 5, , "Incomplete Unicode ticker"
            low = AscW(Mid$(pValue, i + 1, 1)) And &HFFFF&
            If low < &HDC00& Or low > &HDFFF& Then Err.Raise 5, , "Invalid Unicode ticker"
            code = &H10000 + (code - &HD800&) * &H400& + low - &HDC00&
            i = i + 1
        ElseIf code >= &HDC00& And code <= &HDFFF& Then
            Err.Raise 5, , "Invalid Unicode ticker"
        End If
        Select Case code
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95, 126
                result = result & ChrW$(code)
            Case 0 To &H7F&
                result = result & smfMacGoogleHistoryByte(code)
            Case 0 To &H7FF&
                result = result & smfMacGoogleHistoryByte(&HC0& Or (code \ &H40&)) & _
                         smfMacGoogleHistoryByte(&H80& Or (code And &H3F&))
            Case 0 To &HFFFF&
                result = result & smfMacGoogleHistoryByte(&HE0& Or (code \ &H1000&)) & _
                         smfMacGoogleHistoryByte(&H80& Or ((code \ &H40&) And &H3F&)) & _
                         smfMacGoogleHistoryByte(&H80& Or (code And &H3F&))
            Case Else
                result = result & smfMacGoogleHistoryByte(&HF0& Or (code \ &H40000)) & _
                         smfMacGoogleHistoryByte(&H80& Or ((code \ &H1000&) And &H3F&)) & _
                         smfMacGoogleHistoryByte(&H80& Or ((code \ &H40&) And &H3F&)) & _
                         smfMacGoogleHistoryByte(&H80& Or (code And &H3F&))
        End Select
    Next i
    smfMacGoogleHistoryQuery = result
End Function

Private Function smfMacGoogleHistoryByte(ByVal pByte As Long) As String
    smfMacGoogleHistoryByte = "%" & Right$("0" & Hex$(pByte), 2)
End Function
#End If

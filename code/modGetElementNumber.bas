Attribute VB_Name = "modGetElementNumber"
#If Mac Then
Const kVersion = "3.0.2024.07.30-mac.0.7-dev"       ' Mac data repairs and confirmed retired-source guards
#Else
Const kVersion = "3.0.2024.07.30"                   ' Upstream Windows version
#End If
    
Const kElements = 20000                             ' Number of data elements
Dim aParms(1 To kElements) As String                ' Extraction parameters for each element
Dim aFiles(1 To kElements) As String                ' Each element nbr and its "smf-elements-n.txt" file name
Public aConstants(1 To 100) As String               ' Constants for use in RCHGetElementNumber() formulas

Public sElementsLocation As String                  ' Location of element defintion files
Public iInit As Integer                             ' Has data element list been initialized?


Public Function RCHGetElementNumber(ByVal pTicker As String, _
                           Optional ByVal pItem As Integer = 1, _
                           Optional ByVal pError As Variant = "Error", _
                           Optional ByVal pFile As String = "")
                           
    '-----------------------------------------------------------------------------------------------------------*
    ' User defined function to retrieve individual data elements from various web sites
    '-----------------------------------------------------------------------------------------------------------*
    ' 2005.08.01 -- Fix 10-year summary retrieval of stocks with less than 10 years of history
    ' 2005.08.06 -- Simplify code flow and remove redundant modules, add comments
    ' 2005.08.06 -- Add Yahoo! Finance Analyst Estimates page
    ' 2005.08.09 -- Add several Morningstar items
    ' 2005.08.14 -- Add BusinessWeek Online ProSearch Criteria Report elements
    ' 2005.08.17 -- Add BusinessWeek Online Financial Summary page
    '-----------------------------------------------------------------------------------------------> Version 1.0
    ' 2006.02.01 -- Fix MSN "% Growth Rate -- EPS..." items; now "% Growth Rate -- Net Income..."
    ' 2006.02.01 -- Fix YahooKS Cash Flow table extraction tags
    ' 2006.02.01 -- Fix YahooAE "EPS Trends" table extraction
    ' 2006.02.01 -- Fix YahooAE "Revenue Est" table extraction
    ' 2006.02.01 -- Add BarCharts Technical Report page
    ' 2006.02.01 -- Add "Next Earnings Date" from YahooAE page
    '-----------------------------------------------------------------------------------------------> Version 1.1
    ' 2006.02.04 -- Add Earnings.com statistics
    ' 2006.02.06 -- Add StockCharts.com P&F pattern/price objective
    ' 2006.03.07 -- Fix YahooAE "Earnings Growth Rates" table extraction
    ' 2006.03.10 -- Add several MSN Quarterly Income Statement and Balance Sheet items
    ' 2006.03.10 -- Add several Yahoo Quarterly Income Statement and Balance Sheet items
    ' 2006.03.13 -- Add Reuters Quarterly and Annual Income Statements
    ' 2006.03.13 -- Add process to convert extracted numeric data
    ' 2006.03.14 -- Allow multiple table extract labels to be searched for
    ' 2006.03.15 -- Add pFile parameter to RCHGetElementNumber()
    ' 2006.03.15 -- Fix several BWOpin elements that were out of sync with parsing rules database
    ' 2006.04.24 -- Add process to convert strings like 2.34M or 2.34B into numeric values
    ' 2006.05.25 -- Add Google financial statement elements
    ' 2006.05.26 -- Add MSNMoney financial statement elements
    ' 2006.06.18 -- Fixed "short table" data extraction problems
    ' 2006.06.18 -- Modified MSN FYI Advisor extraction
    ' 2006.06.19 -- Modify numeric conversion process to handle strings like "8.52%"
    ' 2006.06.19 -- Modify BWOpin extractions due to web page changes
    ' 2006.06.19 -- Add new dividend items added to YahooKS web page
    ' 2006.06.19 -- Modify MSN Market Capitalization for web page change
    '-----------------------------------------------------------------------------------------------> Version 1.2
    ' 2006.06.21 -- Major rewrite to create internal tables of extraction definitions
    ' 2006.06.26 -- Add "YahooPM" data source for mutual fund performance elements
    ' 2006.06.26 -- Add "YahooPR" data source for mutual fund profile elements
    ' 2006.06.26 -- Add "YahooRK" data source for mutual fund risk-related elements
    ' 2006.06.26 -- Add "YahooHL" data source for mutual fund holdings-related elements
    ' 2006.06.27 -- Add the ability to create a calculated field
    ' 2006.07.01 -- Add the ADVFN annual and quarterly financial statements elements
    ' 2006.07.05 -- Add the Reuters Ratios Comparison page elements
    ' 2006.07.10 -- Add "YahooMS" data source for the Yahoo Market Statistics page (Advances and Declines)
    ' 2006.07.12 -- Add company name elements for "YahooKS" and "MSN" data sources
    ' 2006.07.12 -- Allow for custom "Error" value
    ' 2006.07.14 -- Modify P&F "Bullish"/"Bearish" decision to look at P&F Pattern before price objective
    ' 2006.07.19 -- Fix usage of pFile parameter
    ' 2006.07.27 -- Add "YahooIN" data source for sector and industry numbers and names
    ' 2006.07.27 -- Fix "Zacks" data parsing due to web page changes
    ' 2006.08.04 -- Add "Web Page" value for pTicker parameter
    ' 2006.08.07 -- Force pTicker parameter to be upper case
    ' 2006.08.09 -- Add calculated fields for Pitroski/Altman/Rule One/Magic Formula
    ' 2006.08.10 -- Add 10-year price history fields from Business Week Online
    ' 2006.08.14 -- Add more Morningstar pages/fields
    ' 2006.08.14 -- Fixed Barchart.com Buy/Hold/Sell 3-table cell parsing process
    ' 2006.08.16 -- Added MsgBox with error message to XMLHTTP failure
    ' 2006.08.16 -- Added Google financial statements currenty type and magnitude
    ' 2006.08.19 -- Changed MSN 10-year summary of financial statement items to MSN10 web page source
    ' 2006.08.23 -- Obsoleted a number of MSN Financial Statement elements (MSN changed to Reuters data)
    ' 2006.08.27 -- Added RCHGetTableCell() function
    ' 2006.08.29 -- Added RCHGetWebData() function
    ' 2006.09.02 -- Misc changes to RCHExtractData function related to pCells=0 and pRows
    ' 2006.09.18 -- Removed MsgBox with error message upon XMLHTTP failure
    ' 2006.10.07 -- Add SMFForceRecalculation() macro
    ' 2006.11.14 -- Fix RCHExtractData processing to convert M/B into Million and Trillions on non-numeric data
    ' 2006.12.11 -- Fix several MSN tags related to 5-year average growth rates (Canadian stocks had different tag)
    ' 2006.12.20 -- Change all HTML table header tags to normal table cell tags on web pages retrieved
    ' 2007.01.04 -- Change CCur() usage to CDec() because of precision issues
    ' 2007.01.07 -- Fix Earnings.com extraction of Dividends and Splits data when no Splits table exists
    ' 2007.01.16 -- Modify TickerReset and ArrayReset processing
    ' 2007.01.17 -- Fix RCHGetTableCell() to return vError instead of blanks when not finding new rows
    ' 2007.01.19 -- Fix conversion of "- " and "-- " to return zero values
    ' 2007.01.23 -- Transfer BWHist and BWTech data elements to pick up Telescan data
    ' 2007.01.23 -- Obsolete remaining BWOpin elements (726 to 727)
    ' 2007.01.23 -- Obsolete all BWPro elements (884 to 940)
    ' 2007.01.23 -- Obsolete all BWSumm elements (994 To 1214)
    ' 2007.01.27 -- Fix HTML table header tag processing for RCHGetElementNumber()
    ' 2007.03.19 -- Add "Too many web page retrievals" error message
    ' 2007.03.20 -- Add ability to retrieve RCHGetElementNumber parameters via formula
    ' 2007.05.22 -- Obsolete all Telescan elements since website is gone (728 to 779, 837 to 847, 13891 to 13930)
    '-----------------------------------------------------------------------------------------------> Version 1.3
    ' 2007.06.03 -- Externalize element definitions for RCHGetElementNumber()
    '-----------------------------------------------------------------------------------------------> Version 2.0a
    ' 2007.08.09 -- Added ability to retrieve data via IE object for items that fail with XMLHTTP
    '-----------------------------------------------------------------------------------------------> Version 2.0b
    ' 2007.08.14 -- Add conversion of "---" to return zero values
    '-----------------------------------------------------------------------------------------------> Version 2.0c
    ' 2007.08.23 -- Removed IIF() function using pUseIE
    ' 2007.08.27 -- Added ability to use "~~~~~" in Find1-Find4 external parameters for RCHGetElementNumber()
    '-----------------------------------------------------------------------------------------------> Version 2.0d
    ' 2007.08.30 -- Added smfGetAData and smfGetAParms functions for debugging purposes
    '-----------------------------------------------------------------------------------------------> Version 2.0f
    ' 2007.09.10 -- Changed smfForceRecalculation for EXCEL 2000 processing
    '-----------------------------------------------------------------------------------------------> Version 2.0g
    ' 2007.09.21 -- Restored code to initialize Morningstar for retrieval of data from their web site
    '-----------------------------------------------------------------------------------------------> Version 2.0h
    ' 2007.09.30 -- Added translation of chr(160) to blank
    ' 2007.11.20 -- Broke out major functions into their own modules
    ' 2008.03.16 -- Add P-TYPE parameter to RCHGetElementNumber()
    '-----------------------------------------------------------------------------------------------> Version 2.0i
    ' 2009.12.19 -- Add smfConvertYahooTicker() processing
    ' 2009.12.30 -- Modified P&F element extractions from StockCharts.com
    ' 2010.07.06 -- Fix smfGetAdvFNElementNumber() parameter numbering
    ' 2010.07.18 -- Changed location of smfConvertYahooTicker() processing to prevent "Undefine" error
    ' 2010.10.10 -- Added code to change HTML code &#151; to a normal hyphen
    ' 2010.10.22 -- Added code to change HTML code &mdash; to a normal hyphen
    ' 2010.12.15 -- Added code to return regional setting
    ' 2011.01.26 -- Obsoleted "FYI Alerts" from MSN
    ' 2011.04.27 -- Convert to use smfGetWebPage() function
    ' 2012.01.11 -- Add ability to define element as SMF formula by using leading "=" on formula
    ' 2012.05.13 -- Change placement of smfConvertYahooTicker() for EVALUATE() functions
    ' 2013.08.14 -- Add sAdvFNPrefix to "Version" output
    ' 2014.03.08 -- Add sElementsLocation to allow element definitions on Internet
    ' 2014.03.15 -- Add vTemp processing for returned array from evaluate() results
    ' 2015.04.29 -- Add "Definition" value for pTicker parameter
    ' 2015.09.15 -- Fix "Rule #1 MOS Price" (used to be based on obsoleted MSN data elements)
    ' 2016.05.18 -- Add Application.PathSeparator to ease transition between operating systems
    ' 2016.06.14 -- Add Operating System and Version of computer to "Version" string
    ' 2017.05.05 -- Add smfSetElementsLocation() sub
    ' 2017.05.19 -- Add aConstants() processing
    ' 2017.05.21 -- Add smfGetaConstants() function
    ' 2017.07.23 -- Remove iMorningStar variable
    ' 2018.08.27 -- ERASE aData() array rather than resetting the individual items
    ' 2020.03.09 -- Replace EVALUATE() function with smfEvaluateTwice(), a fix for Microsoft changes
    ' 2023-01-22 -- Mel Pryor (ClimberMel@gmail.com)
    '               Changed version number to reflect updates to other modules
    ' 2023-02-08 -- Version 3.0.2023.02.09 to reflect major overhaul and fix to smfGetYahooHistory
    ' 2023-02-08 -- Version 3.0.2023.02.15 fix smfGetYahooHistory bringing back data as string
    '               Fix data order and sort/resort to match original order
    ' 2023-02-21 -- Fix Issues#9, Issues#19, Issues#18
    '               New version number 3.0.2023.02.24
    ' 2023-03-18 -- Fix for Issue#24 in modGetYahooHistory
    ' 2023-05-08 -- Fix for Issue#27 in smfGetYahooPortfolio_.bas
    '               New version number 3.0.2023.05.09
    ' 2023-05-20 -- Fix Issues #31, 32 New version number 3.0.2023.05.20
    ' 2023-06-05 -- Fix Issue #37, add new Function to smfUtilities_ and update elements-1, 6 & 9
    ' 2023-06-08 -- Fix smfGetYahooPortfolioView
    ' 2023-09-20 -- Issue #56 Fix for smfGetYahooJSONField to work with Excel 64bit
    '               Issue #54 updates to smf-elements9.txt
    '               Updates to smf-elements2.txt
    ' 2024-02-05 -- Updates for Issue #58
    '                 modGetOptionExpirations.bas
    '                 modGetOptionStrikes.bas
    '                 modGetAdvFNElement.bas
    ' 2024-02-05 -- Updates for Issues #66, 67 68 regarding Yahoo change now requires crumb for options functions
    '                 modGetOptionExpirations.bas
    '                 modGetOptionStrikes.bas
    '                 modGetElementNumber.bas (this file to change version)
    '                 elements2.txt and various web pages
    ' 2024-06-09 -- Add "ELEMENTFILE" value for pTicker parameter
    ' 2024-06-10 -- Update Add-in web site name in LoadElementsFromInternet() subroutine
    ' 2024-07-29 -- Create new package with changes from Issue #72, 76, 77, 78, 79, 81, 82, 83, 89, 90
    '                 New add-in: RCH_Stock_Market_Functions-3.0.2024.07.30
    '                 modGetElementNumber.bas
    '                 smfUtilities_.bas
    '                 modGetYahooJSONData.bas
    '                 modGetEconData.bas
    '                 element files 1, 2, 22
    '-----------------------------------------------------------------------------------------------------------*
    ' > Example of an invocation to get The "Trend Spotter" value for IBM from the BarChart website:
    '
    '   =RCHGetElementNumber("IBM", 701)
    '   =RCHGetTableCell("http://quote.barchart.com/texpert.asp?sym=IBM",1,"Trend Spotter",,,,,1,3)
    '
    '   The first is the preferred method.
    '-----------------------------------------------------------------------------------------------------------*
       
    Dim sURL As String
    Dim sTicker1 As String
    Dim sTicker2 As String
    Dim vTemp As Variant
#If Mac Then
    Dim unsupportedReason As String
#End If

    On Error GoTo ErrorExit
    vError = pError
    
    '--------------------------------> Special cases to return immediately
    sTicker1 = UCase(pTicker)
    Select Case True
       Case sTicker1 = "": GoTo ErrorExit
       Case sTicker1 = "NONE": GoTo ErrorExit
       Case sTicker1 = "ERROR": GoTo ErrorExit
       Case sTicker1 = "VERSION": RCHGetElementNumber = "Stock Market Functions add-in, Version " & kVersion & _
                       " (" & ThisWorkbook.Path & "; " _
                            & Application.OperatingSystem & "; " _
                            & Application.Version & "; " _
                            & sAdvFNPrefix & "; " _
                            & sElementsLocation & "; " _
                            & Application.International(xlCountrySetting) & ")"
       Case sTicker1 = "COUNTRY": RCHGetElementNumber = Application.International(xlCountrySetting)
       Case pItem > kElements: RCHGetElementNumber = "EOL"
       Case Else: RCHGetElementNumber = ""
       End Select
    If RCHGetElementNumber <> "" Then Exit Function
    
    '--------------------------------> Load extraction definitions if needed
    If iInit = 0 Then
       iInit = 1
       Erase aData, aParms, aFiles, aConstants
       'For i1 = 1 To kPages
       '    aData(i1, 1) = ""  ' Reset stored ticker array
       '    Next i1
       Call LoadElementsLocation
       For i1 = 0 To 19
           Select Case sElementsLocation
              Case "Internet": Call LoadElementsFromInternet(i1)
              Case Else: Call LoadElementsFromFile(i1)
              End Select
           Next i1
       Call LoadElementsFromFile(20)  ' User Custom Element Definitions
       Call LoadElementsFromFile(21)  ' User default settings
       Call LoadElementsFromFile(22)  ' Constant values
       End If
    
    '--------------------------------> Additional special cases to return immediately
    aParm = Split(aParms(pItem) & ";N/A;N/A;N/A;N/A;N/A;N/A;N/A;N/A;N/A;N/A", ";")
    sTicker2 = smfConvertYahooTicker(sTicker1, aParm(0))
    Select Case True
       Case aParms(pItem) = "": RCHGetElementNumber = "Undefined"
       Case sTicker1 = "DEFINITION": RCHGetElementNumber = aParms(pItem): Exit Function
       Case sTicker1 = "SOURCE": RCHGetElementNumber = aParm(0): Exit Function
       Case sTicker1 = "ELEMENT": RCHGetElementNumber = aParm(1): Exit Function
       Case sTicker1 = "ELEMENTFILE": RCHGetElementNumber = aFiles(pItem): Exit Function
       Case sTicker1 = "WEB PAGE": RCHGetElementNumber = aParm(2): Exit Function
       Case sTicker1 = "P-URL": RCHGetElementNumber = aParm(2): Exit Function
       Case sTicker1 = "P-CELLS": RCHGetElementNumber = aParm(3): Exit Function
       Case sTicker1 = "P-FIND1": RCHGetElementNumber = aParm(4): Exit Function
       Case sTicker1 = "P-FIND2": RCHGetElementNumber = aParm(5): Exit Function
       Case sTicker1 = "P-FIND3": RCHGetElementNumber = aParm(6): Exit Function
       Case sTicker1 = "P-FIND4": RCHGetElementNumber = aParm(7): Exit Function
       Case sTicker1 = "P-ROWS": RCHGetElementNumber = aParm(8): Exit Function
       Case sTicker1 = "P-END": RCHGetElementNumber = aParm(9): Exit Function
       Case sTicker1 = "P-LOOK": RCHGetElementNumber = aParm(10): Exit Function
       Case sTicker1 = "P-TYPE": RCHGetElementNumber = aParm(11): Exit Function
#If Mac Then
       Case smfLiteralUnsupportedReason(CStr(aParm(2)), unsupportedReason)
            RCHGetElementNumber = CVErr(xlErrNA)
            Exit Function
#End If
       Case aParm(0) = "AdvFN-A" And Left(aParm(2), 1) <> "="
            RCHGetElementNumber = smfGetADVFNElement(sTicker1, "A", aParm(3), aParm(4), aParm(5), pError)
       Case aParm(0) = "AdvFN-Q" And Left(aParm(2), 1) <> "="
            RCHGetElementNumber = smfGetADVFNElement(sTicker1, "Q", aParm(3), aParm(4), aParm(5), pError)
       Case aParm(0) = "Evaluate"
#If Mac Then
            If Not smfMacTryTagContentRecipe(Replace(aParm(2), "~~~~~", sTicker2), vTemp) Then
                If Not smfMacTryTableCellRecipe(Replace(aParm(2), "~~~~~", sTicker2), vTemp) Then
                    If smfMacTryNestedRecipe(Replace(aParm(2), "~~~~~", sTicker2), vTemp, pError) Then
                        RCHGetElementNumber = vTemp
                        Exit Function
                    End If
                    ' Unsupported table calls must not re-enter Excel's evaluator.
                    If MacRecipeCallsTableCell(Replace(aParm(2), "~~~~~", sTicker2)) Then
                        RCHGetElementNumber = pError
                        Exit Function
                    End If
                    vTemp = Evaluate("=" & Replace(aParm(2), "~~~~~", sTicker2))
                End If
            End If
            RCHGetElementNumber = vTemp
            Exit Function
#Else
            RCHGetElementNumber = smfEvaluateTwice(Replace(aParm(2), "~~~~~", sTicker2))
#End If
       Case Left(aParm(2), 1) = "="
#If Mac Then
            If Not smfMacTryTagContentRecipe(Replace(aParm(2), "~~~~~", sTicker2), vTemp) Then
                If Not smfMacTryTableCellRecipe(Replace(aParm(2), "~~~~~", sTicker2), vTemp) Then
                    If smfMacTryNestedRecipe(Replace(aParm(2), "~~~~~", sTicker2), vTemp, pError) Then
                        RCHGetElementNumber = vTemp
                        Exit Function
                    End If
                    ' Unsupported table calls must not re-enter Excel's evaluator.
                    If MacRecipeCallsTableCell(Replace(aParm(2), "~~~~~", sTicker2)) Then
                        RCHGetElementNumber = pError
                        Exit Function
                    End If
                    vTemp = Evaluate("=" & Replace(Mid(aParm(2), 2), "~~~~~", sTicker2))
                End If
            End If
#Else
            vTemp = smfEvaluateTwice(Replace(Mid(aParm(2), 2), "~~~~~", sTicker2))
#End If
            If VarType(vTemp) > 8192 Then
               RCHGetElementNumber = vTemp(1)
            Else
               RCHGetElementNumber = vTemp
               End If
#If Mac Then
            Exit Function
#End If
       Case Else: RCHGetElementNumber = ""
       End Select
    If RCHGetElementNumber <> "" Then Exit Function
    '--------------------------------> Preprocess web page data
    sURL = Replace(aParm(2), "~~~~~", sTicker2)
    If aParm(0) = "Calculated" Then
       sData(2) = ""
    Else
       sData(2) = smfGetWebPage(sURL, aParm(11), 0)
       End If
    sData(3) = UCase(sData(2))
    '--------------------------------> Return requested item
    If pFile <> "" Then
       On Error GoTo ErrorExit
       Set wb = Application.Workbooks(pFile).Worksheets(aParm(0))
       iRow = 0
       On Error Resume Next
       iRow = Application.WorksheetFunction.Match(pTicker, wb.Columns("A:A"), 0)
       On Error GoTo ErrorExit
       If iRow > 0 Then
          iCol = 0
          On Error Resume Next
          iCol = Application.WorksheetFunction.Match(pItem, wb.Rows("1:1"), 0)
          On Error GoTo ErrorExit
          If iCol = 0 Then
             Set wb = Application.Workbooks(pFile).Worksheets(aParm(0) & "_2")
             iCol = 0
             On Error Resume Next
             iCol = Application.WorksheetFunction.Match(pItem, wb.Rows("1:1"), 0)
             On Error GoTo ErrorExit
             End If
          If iCol = 0 Then
             Set wb = Application.Workbooks(pFile).Worksheets(aParm(0) & "_3")
             iCol = 0
             On Error Resume Next
             iCol = Application.WorksheetFunction.Match(pItem, wb.Rows("1:1"), 0)
             On Error GoTo ErrorExit
             End If
          RCHGetElementNumber = wb.Cells(iRow, iCol)
          Exit Function
          End If
       End If
    Select Case True
       Case aParm(2) = "?": RCHGetElementNumber = RCHSpecialExtraction("" & aParm(1), sTicker1)
       Case aParm(3) = "?": RCHGetElementNumber = RCHSpecialExtraction("" & aParm(1), sTicker1)
       Case Left(aParm(2), 8) = "Obsolete": RCHGetElementNumber = aParm(2)
       Case Else: RCHGetElementNumber = RCHExtractData(aParm(0), aParm(1), _
                                                       Replace(aParm(4), "~~~~~", sTicker2), _
                                                       Replace(aParm(5), "~~~~~", sTicker2), _
                                                       Replace(aParm(6), "~~~~~", sTicker2), _
                                                       Replace(aParm(7), "~~~~~", sTicker2), _
                                                       aParm(8), aParm(9), aParm(3), aParm(10))
       End Select
    
    Exit Function
ErrorExit: RCHGetElementNumber = vError
    End Function
Private Function RCHSpecialExtraction(pLookup As String, Optional pTicker As String = "")
    On Error GoTo ErrorExit
    Select Case pLookup
        Case "Financial Statements Currency Magnitude" ' Google
             iPos2 = InStr(sData(3), "(EXCEPT FOR PER SHARE ITEMS)")
             iPos2 = InStrRev(sData(3), " OF ", iPos2)
             iPos1 = InStrRev(sData(3), ">IN ", iPos2)
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 4, iPos2 - iPos1 - 4)
        Case "Financial Statements Currency Type" ' Google
             iPos2 = InStr(sData(3), "(EXCEPT FOR PER SHARE ITEMS)")
             iPos1 = InStrRev(sData(3), " OF ", iPos2)
             iPos2 = InStr(iPos1, sData(3), "<")
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 4, iPos2 - iPos1 - 4)
        Case "FYI Alerts" ' MSN
             RCHSpecialExtraction = "No longer available"
             'iPos1 = InStr(sData(3), ">ADVISOR FYI<")
        Case "Company Description" ' MSN
             iPos1 = InStr(sData(3), "<BODY")
             iPos1 = InStr(iPos1, sData(3), "COMPANY REPORT")
             iPos1 = InStr(iPos1, sData(3), "<P>") + 2
             iPos2 = InStr(iPos1, sData(3), "</P>")
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 1, iPos2 - iPos1 - 1)
        Case "Company Name" ' MSN
             iPos1 = InStr(sData(3), "<TITLE>")
             iPos2 = InStr(iPos1, sData(3), " REPORT - ")
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 7, iPos2 - iPos1 - 7)
        Case "Risk Grade"  ' MSN
             iPos1 = InStr(sData(3), "RISK:") + 6
             If iPos1 = 6 Then
                RCHSpecialExtraction = vError
             Else
                i1 = CInt(Mid(sData(2), iPos1, 1))
                RCHSpecialExtraction = Mid("ABCDF", i1, 1)
                End If
        Case "Return Grade"  ' MSN
             iPos1 = InStr(sData(3), "RETURN:") + 8
             If iPos1 = 8 Then
                RCHSpecialExtraction = vError
             Else
                i1 = CInt(Mid(sData(2), iPos1, 1))
                RCHSpecialExtraction = Mid("FDCBA", i1, 1)
                End If
        Case "Quick Summary"  ' MSN
             RCHSpecialExtraction = ""
             iPos1 = InStr(sData(3), "RETURN:") + 8
             iPos2 = InStr(sData(3), "QUICK SUMMARY")
             For i1 = 1 To 20
                 iPos2 = InStr(iPos2, sData(3), "<DD>") + 4
                 If iPos2 > iPos1 Or iPos2 = 4 Then Exit For
                 iPos3 = InStr(iPos2, sData(3), "<B>")
                 iPos4 = InStr(iPos2, sData(3), "</B>")
                 s1 = Mid(sData(2), iPos3 + 3, iPos4 - iPos3 - 3)
                 s2 = Mid(sData(2), iPos2, iPos3 - iPos2)
                 RCHSpecialExtraction = RCHSpecialExtraction & s1 & " -- " & s2 & vbLf
                 Next i1
        Case "StockScouter Rating -- Summary"  ' MSN
             iPos1 = InStr(sData(3), "ALT=""STOCKSCOUTER RATING: ")
             iPos1 = InStr(iPos1, sData(3), "<P>") + 3
             iPos2 = InStr(iPos1, sData(3), "</P>")
             RCHSpecialExtraction = Replace(Replace(Mid(sData(2), iPos1, iPos2 - iPos1), "<b>", ""), "</b>", "")
        Case "Short Term Outlook"  ' MSN
             iPos1 = InStr(sData(3), "SHORT-TERM OUTLOOK")
             iPos1 = InStr(iPos1, sData(3), "<P>") + 3
             iPos2 = InStr(iPos1, sData(3), "</P>")
             RCHSpecialExtraction = Replace(Replace(Mid(sData(2), iPos1, iPos2 - iPos1), "<b>", ""), "</b>", "")
        Case "StockScouter Rating -- Current"  ' MSN
             iPos1 = InStr(sData(3), "ALT=""STOCKSCOUTER RATING: ")
             iPos2 = InStr(iPos1, sData(3), ":") + 2
             iPos3 = InStr(iPos2, sData(3), """")
             RCHSpecialExtraction = CInt(Mid(sData(2), iPos2, iPos3 - iPos2))
        Case "Risk Alert Level" ' Reuters
             iPos1 = InStr(sData(3), "IMAGES/SELLALERT")
             iPos1 = InStr(iPos1, sData(3), "ALT=""") + 5
             iPos2 = InStr(iPos1, sData(3), """")
             RCHSpecialExtraction = Mid(sData(2), iPos1, iPos2 - iPos1)
        Case "P&F -- Pattern" ' Stockcharts
             iPos1 = InStr(sData(3), "P&F PATTERN:")
             If iPos1 = 0 Then
                RCHSpecialExtraction = "No P&F Pattern Found"
                Exit Function
                End If
             iPos2 = InStr(iPos1, sData(3), "</DIV")
             iPos1 = InStrRev(sData(3), ">", iPos2) + 1
             iPos3 = InStrRev(sData(3), "#00AA00", iPos1)
             If iPos1 - iPos3 < 40 Then
                RCHSpecialExtraction = "Bullish -- " & Trim(Mid(sData(2), iPos1, iPos2 - iPos1))
                Exit Function
                End If
             iPos3 = InStrRev(sData(3), "#FF0000", iPos1)
             If iPos1 - iPos3 < 40 Then
                RCHSpecialExtraction = "Bearish -- " & Trim(Mid(sData(2), iPos1, iPos2 - iPos1))
                Exit Function
                End If
             RCHSpecialExtraction = "Unknown -- " & Trim(Mid(sData(2), iPos1, iPos2 - iPos1))
        Case "P&F -- Price Objective"  ' Stockcharts
             iPos1 = InStr(sData(3), " PRICE OBJ. ")
             If iPos1 = 0 Then GoTo ErrorExit
             iPos1 = InStr(iPos1, sData(3), ":") + 2
             iPos2 = InStr(iPos1, sData(3), "<")
             RCHSpecialExtraction = Trim(Mid(sData(2), iPos1, iPos2 - iPos1))
        Case "P&F -- Trend"  ' Stockcharts
             iPos1 = InStr(sData(3), " PRICE OBJ. ")
             If iPos1 > 0 Then
                RCHSpecialExtraction = Mid(sData(2), iPos1 - 7, 7)
             Else
                RCHSpecialExtraction = "Unknown"
                End If
        Case "Next Earnings Date" ' Yahoo
             iPos1 = InStr(sData(3), "NEXT EARNINGS DATE: ") + 20
             If iPos1 = 20 Then
                RCHSpecialExtraction = "N/A"
                Exit Function
                End If
             iPos2 = InStr(iPos1, sData(3), " - ")
             If iPos2 = 0 Then GoTo ErrorExit
             RCHSpecialExtraction = Mid(sData(2), iPos1, iPos2 - iPos1)
        Case "Sector Number" ' Yahoo
             iPos1 = InStr(sData(3), "HTTP://BIZ.YAHOO.COM/P/")
             If iPos1 = 0 Then GoTo ErrorExit
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 23, 1)
        Case "Industry Number" ' Yahoo
             iPos1 = InStr(sData(3), ">INDUSTRY:<")
             iPos1 = InStr(iPos1, sData(3), "HTTP://BIZ.YAHOO.COM/IC/")
             If iPos1 = 0 Then GoTo ErrorExit
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 24, 3)
        Case "Industry Symbol" ' Yahoo
             iPos1 = InStr(sData(3), ">^")
             If iPos1 = 0 Then GoTo ErrorExit
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 1, 8)
        Case "Company Name" ' Yahoo
             iPos2 = InStr(sData(3), " (" & pTicker & ")</B>")
             iPos1 = InStrRev(sData(3), "<B>", iPos2)
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 3, iPos2 - iPos1 - 3)
        Case "Fund Profile -- Morningstar Rating" ' Yahoo
             Select Case True
                Case InStr(sData(3), "/STAR1.GIF") > 0: RCHSpecialExtraction = 1
                Case InStr(sData(3), "/STAR2.GIF") > 0: RCHSpecialExtraction = 2
                Case InStr(sData(3), "/STAR3.GIF") > 0: RCHSpecialExtraction = 3
                Case InStr(sData(3), "/STAR4.GIF") > 0: RCHSpecialExtraction = 4
                Case InStr(sData(3), "/STAR5.GIF") > 0: RCHSpecialExtraction = 5
                Case Else: RCHSpecialExtraction = vError
                End Select
        Case "Fund Profile -- Last Dividend -- Date" ' Yahoo
             iPos1 = InStr(sData(3), ">FUND OPERATIONS")
             If iPos1 = 0 Then GoTo ErrorExit
             iPos1 = InStr(iPos1, sData(3), "LAST DIVIDEND")
             If iPos1 = 0 Then GoTo ErrorExit
             iPos1 = InStr(iPos1, sData(3), "(")
             If iPos1 = 0 Then GoTo ErrorExit
             iPos2 = InStr(iPos1, sData(3), ")")
             If iPos2 < iPos1 Then GoTo ErrorExit
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 1, iPos2 - iPos1 - 1)
        Case "Fund Profile -- Last Cap Gain -- Date" ' Yahoo
             iPos1 = InStr(sData(3), ">FUND OPERATIONS")
             If iPos1 = 0 Then GoTo ErrorExit
             iPos1 = InStr(iPos1, sData(3), "LAST CAP GAIN")
             If iPos1 = 0 Then GoTo ErrorExit
             iPos1 = InStr(iPos1, sData(3), "(")
             If iPos1 = 0 Then GoTo ErrorExit
             iPos2 = InStr(iPos1, sData(3), ")")
             If iPos2 < iPos1 Then GoTo ErrorExit
             RCHSpecialExtraction = Mid(sData(2), iPos1 + 1, iPos2 - iPos1 - 1)
        Case "Piotroski 1 (Positive Net Income)"
             n1 = RCHGetElementNumber(pTicker, 8806)  ' FQ1, Net Income (Continuing Operations)
             n2 = RCHGetElementNumber(pTicker, 8807)  ' FQ2, Net Income (Continuing Operations)
             n3 = RCHGetElementNumber(pTicker, 8808)  ' FQ3, Net Income (Continuing Operations)
             n4 = RCHGetElementNumber(pTicker, 8809)  ' FQ4, Net Income (Continuing Operations)
             RCHSpecialExtraction = IIf((n1 + n2 + n3 + n4) > 0, 1, 0)
        Case "Piotroski 2 (Positive Operating Cash Flow)"
             n1 = RCHGetElementNumber(pTicker, 11326) ' FQ1, YTD Net Cash Flow (Continuing Operations)
             n2 = RCHGetElementNumber(pTicker, 11330) ' FQ5, YTD Net Cash Flow (Continuing Operations)
             n3 = RCHGetElementNumber(pTicker, 6856)  ' FY1, Net Cash Flow (Continuing Operations)
             RCHSpecialExtraction = IIf(n1 - n2 + n3 > 0, 1, 0)
        Case "Piotroski 3 (Increasing Net Income)"
             n1 = RCHGetElementNumber(pTicker, 8066)     ' FQ1, Ending Quarter
             If n1 = 4 Then
                n6 = RCHGetElementNumber(pTicker, 5596)  ' FY1, Net Income (Continuing Operations)
                N7 = RCHGetElementNumber(pTicker, 5597)  ' FY2, Net Income (Continuing Operations)
             Else
                n2 = RCHGetElementNumber(pTicker, 8806)  ' FQ1, Net Income (Continuing Operations)
                n3 = RCHGetElementNumber(pTicker, 8807)  ' FQ2, Net Income (Continuing Operations)
                n4 = RCHGetElementNumber(pTicker, 8808)  ' FQ3, Net Income (Continuing Operations)
                n5 = RCHGetElementNumber(pTicker, 8809)  ' FQ4, Net Income (Continuing Operations)
                n6 = n2 + n3 + n4 + n5
                N7 = RCHGetElementNumber(pTicker, 5596)  ' FY1, Net Income (Continuing Operations)
                End If
             RCHSpecialExtraction = IIf(n6 > N7, 1, 0)
        Case "Piotroski 4 (Operating Cash flow exceeds Net Income)"
             n1 = RCHGetElementNumber(pTicker, 11326) ' FQ1, YTD Net Cash Flow (Continuing Operations)
             n2 = RCHGetElementNumber(pTicker, 11330) ' FQ5, YTD Net Cash Flow (Continuing Operations)
             n3 = RCHGetElementNumber(pTicker, 6856)  ' FY1, Net Cash Flow (Continuing Operations)
             n4 = RCHGetElementNumber(pTicker, 8806)  ' FQ1, Net Income (Continuing Operations)
             n5 = RCHGetElementNumber(pTicker, 8807)  ' FQ2, Net Income (Continuing Operations)
             n6 = RCHGetElementNumber(pTicker, 8808)  ' FQ3, Net Income (Continuing Operations)
             N7 = RCHGetElementNumber(pTicker, 8809)  ' FQ4, Net Income (Continuing Operations)
             RCHSpecialExtraction = IIf(n1 - n2 + n3 > n4 + n5 + n6 + N7, 1, 0)
        Case "Piotroski 5 (Decreasing ratio of long-term debt to assets )"
             n1 = RCHGetElementNumber(pTicker, 8066)     ' FQ1, Ending Quarter
             If n1 = 4 Then
                n2 = RCHGetElementNumber(pTicker, 6376)  ' FY1, Long Term Debt
                n3 = RCHGetElementNumber(pTicker, 6266)  ' FY1, Total Assets
                n4 = RCHGetElementNumber(pTicker, 6377)  ' FY2, Long Term Debt
                n5 = RCHGetElementNumber(pTicker, 6267)  ' FY2, Total Assets
             Else
                n2 = RCHGetElementNumber(pTicker, 10366) ' FQ1, Long Term Debt
                n3 = RCHGetElementNumber(pTicker, 10146) ' FQ1, Total Assets
                n4 = RCHGetElementNumber(pTicker, 6376)  ' FY1, Long Term Debt
                n5 = RCHGetElementNumber(pTicker, 6266)  ' FY1, Total Assets
                End If
             RCHSpecialExtraction = IIf((n2 / n3) < (n4 / n5), 1, 0)
        Case "Piotroski 6 (Increasing Current Ratio)"
             n1 = RCHGetElementNumber(pTicker, 8066)     ' FQ1, Ending Quarter
             If n1 = 4 Then
                n2 = RCHGetElementNumber(pTicker, 6116)  ' FY1, Current Assets
                n3 = RCHGetElementNumber(pTicker, 6366)  ' FY1, Current Liabilities
                n4 = RCHGetElementNumber(pTicker, 6117)  ' FY2, Current Assets
                n5 = RCHGetElementNumber(pTicker, 6367)  ' FY2, Current Liabilities
             Else
                n2 = RCHGetElementNumber(pTicker, 9846)  ' FQ1, Current Assets
                n3 = RCHGetElementNumber(pTicker, 10346) ' FQ1, Current Liabilities
                n4 = RCHGetElementNumber(pTicker, 6116)  ' FY1, Current Assets
                n5 = RCHGetElementNumber(pTicker, 6366)  ' FY1, Current Liabilities
                End If
             RCHSpecialExtraction = IIf((n2 / n3) > (n4 / n5), 1, 0)
        Case "Piotroski 7 (No increase in outstanding shares)"
             n1 = RCHGetElementNumber(pTicker, 8066)     ' FQ1, Ending Quarter
             If n1 = 4 Then
                n2 = RCHGetElementNumber(pTicker, 6646)  ' FY1, Total Common Shares Out
                n3 = RCHGetElementNumber(pTicker, 6647)  ' FY2, Total Common Shares Out
             Else
                n2 = RCHGetElementNumber(pTicker, 10906) ' FQ1, Total Common Shares Out
                n3 = RCHGetElementNumber(pTicker, 6646)  ' FY1, Total Common Shares Out
                End If
             RCHSpecialExtraction = IIf(n2 > n3, 0, 1)
        Case "Piotroski 8 (Increasing Gross Margins)"
             n1 = RCHGetElementNumber(pTicker, 8066)     ' FQ1, Ending Quarter
             If n1 = 4 Then
                n6 = RCHGetElementNumber(pTicker, 5346)  ' FY1, Gross Operating Profit
                N7 = RCHGetElementNumber(pTicker, 5347)  ' FY2, Gross Operating Profit
                n8 = RCHGetElementNumber(pTicker, 5286)  ' FY1, Operating Revenue
                n9 = RCHGetElementNumber(pTicker, 5287)  ' FY2, Operating Revenue
             Else
                n2 = RCHGetElementNumber(pTicker, 8306)  ' FQ1, Gross Operating Profit
                n3 = RCHGetElementNumber(pTicker, 8307)  ' FQ2, Gross Operating Profit
                n4 = RCHGetElementNumber(pTicker, 8308)  ' FQ3, Gross Operating Profit
                n5 = RCHGetElementNumber(pTicker, 8309)  ' FQ4, Gross Operating Profit
                n6 = n2 + n3 + n4 + n5
                N7 = RCHGetElementNumber(pTicker, 5346)  ' FY1, Gross Operating Profit
                n2 = RCHGetElementNumber(pTicker, 8186)  ' FQ1, Operating Revenue
                n3 = RCHGetElementNumber(pTicker, 8187)  ' FQ2, Operating Revenue
                n4 = RCHGetElementNumber(pTicker, 8188)  ' FQ3, Operating Revenue
                n5 = RCHGetElementNumber(pTicker, 8189)  ' FQ4, Operating Revenue
                n8 = n2 + n3 + n4 + n5
                n9 = RCHGetElementNumber(pTicker, 5286)  ' FY1, Operating Revenue
                End If
             RCHSpecialExtraction = IIf((n6 / n8) > (N7 / n9), 1, 0)
        Case "Piotroski 9 (Increasing Asset Turnover)"
             n1 = RCHGetElementNumber(pTicker, 8066)     ' FQ1, Ending Quarter
             If n1 = 4 Then
                n6 = RCHGetElementNumber(pTicker, 5286)  ' FY1, Operating Revenue
                N7 = RCHGetElementNumber(pTicker, 6266)  ' FY1, Total Assets
                n8 = RCHGetElementNumber(pTicker, 5287)  ' FY2, Operating Revenue
                n9 = RCHGetElementNumber(pTicker, 6267)  ' FY2, Total Assets
             Else
                n2 = RCHGetElementNumber(pTicker, 8186)  ' FQ1, Operating Revenue
                n3 = RCHGetElementNumber(pTicker, 8187)  ' FQ2, Operating Revenue
                n4 = RCHGetElementNumber(pTicker, 8188)  ' FQ3, Operating Revenue
                n5 = RCHGetElementNumber(pTicker, 8189)  ' FQ4, Operating Revenue
                n6 = n2 + n3 + n4 + n5
                N7 = RCHGetElementNumber(pTicker, 10146) ' FQ1, Total Assets
                n8 = RCHGetElementNumber(pTicker, 5286)  ' FY1, Operating Revenue
                n9 = RCHGetElementNumber(pTicker, 6266)  ' FY1, Total Assets
                End If
             RCHSpecialExtraction = IIf((n6 / N7) > (n8 / n9), 1, 0)
        Case "Piotroski F-Score"
             n1 = RCHGetElementNumber(pTicker, 15001)
             n2 = RCHGetElementNumber(pTicker, 15002)
             n3 = RCHGetElementNumber(pTicker, 15003)
             n4 = RCHGetElementNumber(pTicker, 15004)
             n5 = RCHGetElementNumber(pTicker, 15005)
             n6 = RCHGetElementNumber(pTicker, 15006)
             N7 = RCHGetElementNumber(pTicker, 15007)
             n8 = RCHGetElementNumber(pTicker, 15008)
             n9 = RCHGetElementNumber(pTicker, 15009)
             RCHSpecialExtraction = n1 + n2 + n3 + n4 + n5 + n6 + N7 + n8 + n9
        Case "Altman Z-Score"
             n1 = RCHGetElementNumber(pTicker, 10786)  ' FQ1, Working Capital
             n2 = RCHGetElementNumber(pTicker, 10146)  ' FQ1, Total Assets
             n3 = RCHGetElementNumber(pTicker, 10646)  ' FQ1, Retained Earnings
             n4 = RCHGetElementNumber(pTicker, 8666)   ' FQ1, EBIT
             n5 = RCHGetElementNumber(pTicker, 8667)   ' FQ2, EBIT
             n6 = RCHGetElementNumber(pTicker, 8668)   ' FQ3, EBIT
             N7 = RCHGetElementNumber(pTicker, 8669)   ' FQ4, EBIT
             n8 = n4 + n5 + n6 + N7
             n9 = RCHGetElementNumber(pTicker, 941)   ' Market Capitalization
             n10 = RCHGetElementNumber(pTicker, 10526) ' Total Liabilities
             n11 = RCHGetElementNumber(pTicker, 8186)  ' FQ1, Operating Revenue
             n12 = RCHGetElementNumber(pTicker, 8187)  ' FQ2, Operating Revenue
             n13 = RCHGetElementNumber(pTicker, 8188)  ' FQ3, Operating Revenue
             n14 = RCHGetElementNumber(pTicker, 8189)  ' FQ4, Operating Revenue
             n15 = n11 + n12 + n13 + n14
             RCHSpecialExtraction = 1.2 * (n1 / n2) + 1.4 * (n3 / n2) + 3.3 * (n8 / n2) + 0.6 * (n9 / n10 / 1000) + (n15 / n2)
        Case "Rule #1 MOS Price"
             n1 = RCHGetElementNumber(pTicker, 13630)  ' 5-Year High P/E from Reuter's
             n2 = RCHGetElementNumber(pTicker, 13634)  ' 5-Year Low P/E from Reuter's
             n3 = RCHGetElementNumber(pTicker, 962)    ' Current EPS from Yahoo
             n4 = RCHGetElementNumber(pTicker, 621)    ' 5-Year Projected Growth Rate from Yahoo
             If n1 > 50 Then n1 = 50
             n5 = FV(n4, 10, 0, -n3)
             n6 = PV(0.15, 10, 0, -n5 * (n1 + n2) / 2) / 2
             RCHSpecialExtraction = n6
        Case "Magic Formula Investing -- Earnings Yield"
             n1 = RCHGetElementNumber(pTicker, 949)    ' Enterprise value to EBITDA
             RCHSpecialExtraction = 1 / n1
        Case "Magic Formula Investing -- Return on Capital"
             n1 = RCHGetElementNumber(pTicker, 960)    ' EBITDA
             n2 = RCHGetElementNumber(pTicker, 964)    ' Cash
             n3 = RCHGetElementNumber(pTicker, 10026)  ' FQ1, Net Fixed Assets (Plant & Equipment)
             RCHSpecialExtraction = n1 / (n2 + 1000 * n3)
        Case Else: RCHSpecialExtraction = vError
        End Select
        Exit Function
ErrorExit: RCHSpecialExtraction = vError
    End Function

Sub LoadElementsFromFile(pSuffix As Variant)
    '------------------------------------------------------------------------------------------------------*
    ' 2017.05.05 -- Add processing for settings
    ' 2017.05.19 -- Add processing for constants
    ' 2024.06.09 -- Add processing to get the "smf-elements-n.txt" file name
    '               the element was last loaded from.
    '------------------------------------------------------------------------------------------------------*
    On Error GoTo ErrorExit
    Dim sElementFile As String
    
    sElementFile = ThisWorkbook.Path & Application.PathSeparator & "smf-elements-" & pSuffix & ".txt"
#If Mac Then
    LoadMacElementsText smfMacReadElementText(ThisWorkbook.Path, CStr(pSuffix)), sElementFile
#Else
    Open sElementFile For Input As #1
    On Error Resume Next
    
    Do Until EOF(1) = True
       Line Input #1, sLine
       Select Case True
          Case Left(sLine, 1) = "'"
          Case Left(UCase(sLine), 7) = "SETTING"
               Evaluate (smfWord(sLine, 3, ";"))
          Case Left(UCase(sLine), 8) = "CONSTANT"
               i1 = CInt(smfWord(sLine, 2, ";"))
               aConstants(i1) = smfWord(sLine, 4, ";")
          Case Else
               iPos1 = InStr(sLine, ";")
               aParms(CInt(Left(sLine, iPos1 - 1))) = Mid(sLine, iPos1 + 1)
               aFiles(CInt(Left(sLine, iPos1 - 1))) = sElementFile & "   (Date: " & FileDateTime(sElementFile) & ")"
          End Select
       Loop
    Close #1
#End If
ErrorExit:
    End Sub

#If Mac Then
Private Sub LoadMacElementsText(ByVal pText As String, ByVal pFile As String)
    Dim lines As Variant, line As Variant, sLine As String
    Dim separatorAt As Long, item As Long
    If Len(pText) = 0 Then Exit Sub
    pText = Replace(Replace(pText, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(pText, vbLf)
    On Error Resume Next
    For Each line In lines
        sLine = CStr(line)
        Select Case True
            Case Len(Trim$(sLine)) = 0
            Case Left$(sLine, 1) = "'"
            Case Left$(UCase$(sLine), 7) = "SETTING"
                Evaluate smfWord(sLine, 3, ";")
            Case Left$(UCase$(sLine), 8) = "CONSTANT"
                Err.Clear
                item = CInt(smfWord(sLine, 2, ";"))
                If Err.Number = 0 And item >= LBound(aConstants) And item <= UBound(aConstants) Then
                    aConstants(item) = smfWord(sLine, 4, ";")
                End If
            Case Else
                separatorAt = InStr(sLine, ";")
                If separatorAt > 1 Then
                    Err.Clear
                    item = CInt(Left$(sLine, separatorAt - 1))
                    If Err.Number = 0 And item >= LBound(aParms) And item <= UBound(aParms) Then
                        aParms(item) = Mid$(sLine, separatorAt + 1)
                        aFiles(item) = pFile & " (Mac helper)"
                    End If
                End If
        End Select
    Next line
    Err.Clear
End Sub
#End If

Public Function SMFMacElementsSelfTest() As Long
#If Mac Then
    Dim oldParm As String, oldFile As String, oldConstant As String
    Dim errorNumber As Long, errorText As String
    oldParm = aParms(19999)
    oldFile = aFiles(19999)
    oldConstant = aConstants(99)
    On Error GoTo RestoreFailed
    LoadMacElementsText "'Comment" & vbCrLf & vbCrLf & _
        "19999;FinViz;Fixture;=SUM(1,2)" & vbLf & _
        "bad;ignored" & vbCr & _
        "CONSTANT;99;Fixture;value" & vbLf & "CONSTANT;bad;Fixture;wrong", "fixture.txt"
    If aParms(19999) <> "FinViz;Fixture;=SUM(1,2)" Then Err.Raise 5, , "Element definition parsing failed"
    If aFiles(19999) <> "fixture.txt (Mac helper)" Then Err.Raise 5, , "Element filename was lost"
    If aConstants(99) <> "value" Then Err.Raise 5, , "Constant parsing failed"
    SMFMacElementsSelfTest = 3 + smfUnsupportedFieldSelfChecks()
RestoreState:
    aParms(19999) = oldParm
    aFiles(19999) = oldFile
    aConstants(99) = oldConstant
    If errorNumber <> 0 Then Err.Raise errorNumber, "SMFMacElementsSelfTest", errorText
    Exit Function
RestoreFailed:
    errorNumber = Err.Number
    errorText = Err.Description
    Resume RestoreState
#Else
    Err.Raise 5, , "This test requires Excel for Mac"
#End If
End Function

Sub LoadElementsFromInternet(pSuffix As Variant)
    '-----------------------------------------------------------------------------------------------------------*
    ' 2016.05.18 -- Change to RCHGetURLData to ease transition between operating systems
    ' 2024.06.10 -- Update Add-in web site name.
    '               Store file location of element.
    '-----------------------------------------------------------------------------------------------------------*
    On Error GoTo ErrorExit
    Dim sElementFile As String
    
'   sElementFile = "http://ogres-crypt.com/SMF/Elements/smf-elements-" & pSuffix & ".txt"
    sElementFile = "https://climbermel.github.io/SMF_Add-in/Elements/smf-elements-" & pSuffix & ".txt"
    
    s1 = RCHGetURLData(sElementFile)
    If s1 = "Error" Then Exit Sub
    v1 = Split(s1, Chr(13) & Chr(10))
    
    For i1 = 0 To UBound(v1)
       Select Case Left(Trim(v1(i1)) & " ", 1)
          Case "'"
          Case " "
          Case Else
               iPos1 = InStr(v1(i1), ";")
               aParms(CInt(Left(v1(i1), iPos1 - 1))) = Mid(v1(i1), iPos1 + 1)
               aFiles(CInt(Left(v1(i1), iPos1 - 1))) = sElementFile
          End Select
        Next i1
    
ErrorExit:
    End Sub
    
Sub smfSetElementsLocation(p1 As String)
    sElementsLocation = p1
    End Sub
    
Sub LoadElementsLocation()
    If sElementsLocation = "" Then sElementsLocation = "Local"
    On Error GoTo ErrorExit
#If Mac Then
    Dim locationText As String, locationLines As Variant
    locationText = smfMacReadElementLocation(ThisWorkbook.Path)
    If Len(locationText) > 0 Then
        locationLines = Split(Replace(Replace(locationText, vbCrLf, vbLf), vbCr, vbLf), vbLf)
        sElementsLocation = CStr(locationLines(0))
    End If
#Else
    Open ThisWorkbook.Path & Application.PathSeparator & "smf-Elements-Location.txt" For Input As #1
    Line Input #1, sElementsLocation
    Close #1
#End If
ErrorExit:
    End Sub

Public Function smfGetAParms(p1 As Integer)
    smfGetAParms = aParms(p1)
    End Function

Public Function smfGetAConstants(p1 As Integer)
    '-----------------------------------------------------------------------------------------------------------*
    ' 2017.05.21 -- New function to view constant values
    '-----------------------------------------------------------------------------------------------------------*
    smfGetAConstants = aConstants(p1)
    End Function

Public Function smfGetaFiles(p1 As Integer)
    '-----------------------------------------------------------------------------------------------------------*
    ' 2024.06.29 -- New test function to view element's file location.
    '-----------------------------------------------------------------------------------------------------------*
    smfGetaFiles = aFiles(p1)
    End Function


#If Mac Then
Private Function MacRecipeCallsTableCell(ByVal pExpression As String) As Boolean
    Dim i As Long, startAt As Long, nextAt As Long
    Dim quoted As Boolean, character As String, token As String
    i = 1
    Do While i <= Len(pExpression)
        character = Mid$(pExpression, i, 1)
        If character = Chr$(34) Then
            If quoted And i < Len(pExpression) Then
                If Mid$(pExpression, i + 1, 1) = Chr$(34) Then
                    i = i + 2
                    GoTo NextCharacter
                End If
            End If
            quoted = Not quoted
        ElseIf Not quoted Then
            If character Like "[A-Za-z0-9_]" Then
                startAt = i
                Do While i <= Len(pExpression)
                    character = Mid$(pExpression, i, 1)
                    If Not character Like "[A-Za-z0-9_]" Then Exit Do
                    i = i + 1
                Loop
                token = Mid$(pExpression, startAt, i - startAt)
                If StrComp(token, "RCHGetTableCell", vbTextCompare) = 0 Then
                    nextAt = i
                    Do While nextAt <= Len(pExpression)
                        character = Mid$(pExpression, nextAt, 1)
                        If character <> " " And character <> vbTab And character <> vbCr And character <> vbLf Then Exit Do
                        nextAt = nextAt + 1
                    Loop
                    If Mid$(pExpression, nextAt, 1) = "(" Then
                        MacRecipeCallsTableCell = True
                        Exit Function
                    End If
                End If
                GoTo NextCharacter
            End If
        End If
        i = i + 1
NextCharacter:
    Loop
End Function

Public Function SMFMacTableGuardSelfTest() As Long
    Dim expression As Variant, actual As Variant, checks As Long
    Dim oldParm As String, oldInit As Integer, errorNumber As Long, errorText As String
    Dim oldGuardError As Variant
    Dim beforeURL As String, beforeError As String, beforeStatus As Long
    oldParm = aParms(19997)
    oldGuardError = vError
    oldInit = iInit
    beforeURL = SMFMacLastURL()
    beforeError = SMFMacLastError()
    beforeStatus = SMFMacHTTPStatus()
    On Error GoTo Failed
    For Each expression In Array( _
        "=RCHGetTableCell(""x"",1)", _
        "=PROPER(RCHGetTableCell(""x"",1))", _
        "=rchgettablecell (""x"",1)*1000", _
        "=modGetTableCell.RCHGetTableCell(""x"",1)", _
        "=SUM(1," & vbTab & "RCHGetTableCell" & vbTab & "(""x"",1))")
        If Not MacRecipeCallsTableCell(CStr(expression)) Then Err.Raise 5, , "Table call was not detected"
        checks = checks + 1
    Next expression
    For Each expression In Array( _
        "=""RCHGetTableCell(""""x"""",1)""", _
        "=OtherRCHGetTableCell(""x"",1)", _
        "=RCHGetTableCellExtra(""x"",1)", _
        "=RCHGetTableCell", _
        "=RCHGetTableCell + 1", _
        "=SUM(1,2)", _
        "=""first"""" RCHGetTableCell(""""x"""",1)""", _
        "=123RCHGetTableCell(""x"",1)")
        If MacRecipeCallsTableCell(CStr(expression)) Then Err.Raise 5, , "Non-table expression was blocked"
        checks = checks + 1
    Next expression
    iInit = 1
    aParms(19997) = "Fixture;Compound;=PROPER(RCHGetTableCell(""https://example.invalid/no-network"",1))"
    actual = RCHGetElementNumber("TEST", 19997, "unsupported-table")
    If CStr(actual) <> "unsupported-table" Then Err.Raise 5, , "Compound table error changed"
    checks = checks + 1
    actual = RCHGetElementNumber("TEST", 19997, "")
    If VarType(actual) <> vbString Then Err.Raise 5, , "Empty error changed type"
    If CStr(actual) <> "" Then Err.Raise 5, , "Empty error fell through preprocessing"
    checks = checks + 1
    actual = RCHGetElementNumber("TEST", 19997, -77)
    If VarType(actual) <> vbInteger Then Err.Raise 5, , "Numeric error changed type"
    If CInt(actual) <> -77 Then Err.Raise 5, , "Numeric error changed value"
    checks = checks + 1
    aParms(19997) = "Evaluate;Compound;PROPER(RCHGetTableCell(""https://example.invalid/no-network"",1))"
    actual = RCHGetElementNumber("TEST", 19997, "evaluate-blocked")
    If CStr(actual) <> "evaluate-blocked" Then Err.Raise 5, , "Evaluate-source table call was not blocked"
    checks = checks + 1
    If SMFMacLastURL() <> beforeURL Or SMFMacLastError() <> beforeError Or _
       SMFMacHTTPStatus() <> beforeStatus Then Err.Raise 5, , "Guard fixtures made an HTTP request"
    SMFMacTableGuardSelfTest = checks
RestoreState:
    aParms(19997) = oldParm
    vError = oldGuardError
    iInit = oldInit
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, "SMFMacTableGuardSelfTest", errorText
    Exit Function
Failed:
    errorNumber = Err.Number
    errorText = Err.Description
    Resume RestoreState
End Function

' Keep private element-fixture state inside the module that owns aParms.
Public Function SMFMacEmptyRecipeSelfTest() As Long
    Dim recipes As Variant, actual As Variant, i As Long, checks As Long
    Dim oldParm As String, oldInit As Integer, oldError As Variant, fallback As String
    Dim beforeURL As String, beforeError As String, beforeStatus As Long
    Dim errorNumber As Long, errorText As String
    oldParm = aParms(19996): oldInit = iInit: oldError = vError
    beforeURL = SMFMacLastURL(): beforeError = SMFMacLastError(): beforeStatus = SMFMacHTTPStatus()
    On Error GoTo Failed
    recipes = Array( _
        "Fixture;Empty;=smfStrExtr(""abc"",""missing"",""~"")", _
        "Evaluate;Empty;smfStrExtr(""abc"",""missing"",""~"")", _
        "Fixture;EmptyError;=smfConvertData(1E999)", _
        "Evaluate;EmptyError;smfConvertData(1E999)")
    iInit = 1
    For i = LBound(recipes) To UBound(recipes)
        aParms(19996) = CStr(recipes(i))
        fallback = IIf(i < 2, "outer error", "")
        actual = RCHGetElementNumber("TEST", 19996, fallback)
        If VarType(actual) <> vbString Then Err.Raise 5, , "Empty recipe result changed type"
        If CStr(actual) <> "" Then Err.Raise 5, , "Empty recipe fell through to another path"
        checks = checks + 1
    Next i
    If SMFMacLastURL() <> beforeURL Or SMFMacLastError() <> beforeError Or _
       SMFMacHTTPStatus() <> beforeStatus Then Err.Raise 5, , "Empty recipe made an HTTP request"
    checks = checks + 1
    SMFMacEmptyRecipeSelfTest = checks
RestoreState:
    aParms(19996) = oldParm: iInit = oldInit: vError = oldError
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, "SMFMacEmptyRecipeSelfTest", errorText
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Resume RestoreState
End Function
#End If

#If Mac Then
' BEGIN GENERATED UNSUPPORTED SOURCES
' Generated from mac/not_supported.json; run mac/build_not_supported.py --write.
' Registry SHA256: 9086a5e9cf0fccdcf93a7c60302fd5c3e01c39d3755fe4b2d824efb731a539a8
Private Function smfUnsupportedRegistry() As Collection
    Dim rows As Collection
    Set rows = New Collection
    rows.Add _
        Array( _
            "msn-options", _
            "smfGetMSNOptionQuote, smfGetOptionQuotes", _
            "MSN option quotes", _
            "MSN", _
            "Not Supported", _
            "SMF marks the MSN option quote function obsolete.", _
            "Use a supported option source; positive Yahoo bid/ask prices still need verification.", _
            "code/modGetOptionQuotes.bas: MSN obsolete return dated 2014-12-23", _
            "2026-11-09")
    rows.Add _
        Array( _
            "optionsxpress-options", _
            "smfGetOXOptionQuote, smfGetOX2OptionQuote, smfGetOX3OptionQuote, smfGetOptionQuotes, smfGetOptionExpirations, smfGetOptionStrikes", _
            "OX, OX2, OX3 quotes; OX expirations/strikes", _
            "OptionsXpress", _
            "Not Supported", _
            "SMF records that optionsXpress is no longer a valid data source.", _
            "Yahoo expiration/strike modes remain available; quote availability needs separate verification.", _
            "code/modGetOptionQuotes.bas: optionsXpress retirement dated 2017-10-10; code/modGetOptionExpirations.bas: OX dispatch disabled; code/modGetOptionStrikes.bas: retirement comment", _
            "2026-11-09")
    rows.Add _
        Array( _
            "yahoo-quote-csv", _
            "RCHGetYahooQuotes", _
            "Ticker/quote-code mode with nonempty pItems", _
            "Yahoo legacy quote CSV", _
            "Not Supported", _
            "SMF explicitly marks the old ticker/quote-code CSV mode obsolete.", _
            "Keep CSV-URL mode with empty pItems, or use the supported Yahoo JSON/history functions.", _
            "code/modGetYahooQuotes.bas: obsolete-mode notice dated 2023-03-01; docs/Documentation/doc-links/RCHGetYahooQuotes-Function_files/RCHGetYahooQuotes-Function.htm", _
            "2026-11-09")
    rows.Add _
        Array( _
            "yahoo-yql", _
            "smfYahooAPIData", _
            "All YQL feeds A-F", _
            "Yahoo YQL", _
            "Not Supported", _
            "Yahoo retired query.yahooapis.com YQL on 2019-01-03, as recorded in the SMF change log.", _
            "Use supported Yahoo JSON/history functions for the fields they provide.", _
            "code/modYahooAPIData.bas: shutdown notice dated 2023-05-28; docs/Documentation/doc-links/Change-Log_files/Change-Log.htm: 2022-12-11 entry", _
            "2026-11-09")
    rows.Add _
        Array( _
            "google-option-chain", _
            "smfGetGoogleOptionQuote, smfGetOptionQuotes, smfGetOptionExpirations, smfGetOptionStrikes", _
            "G/GOOGLE legacy option chain", _
            "Google legacy option chain", _
            "Not Supported", _
            "The SMF option-strike manual explicitly labels the Google option source obsolete.", _
            "Yahoo expiration/strike modes remain available; positive option prices need separate verification.", _
            "docs/Documentation/doc-links/smfGetOptionStrikes-Function_files/smfGetOptionStrikes-Function.htm: obsolete 8/G/OX modes; docs/Documentation/doc-links/smfGetOptionExpirations_files/smfGetOptionExpirations.htm: Yahoo-only source note", _
            "2026-11-09")
    rows.Add _
        Array( _
            "oic-legacy-options", _
            "smfGet888OptionQuote, smfGetOptionQuotes, smfGetOptionExpirations, smfGetOptionStrikes", _
            "8/888 legacy oic_adv_options.j path", _
            "OIC legacy option chain", _
            "Not Supported", _
            "The SMF option-strike manual explicitly labels source 8 obsolete; these calls use that same legacy endpoint.", _
            "Yahoo expiration/strike modes remain available; other iVolatility services are outside this guard.", _
            "docs/Documentation/doc-links/smfGetOptionStrikes-Function_files/smfGetOptionStrikes-Function.htm: obsolete 8/G/OX modes; code/modGetOptionQuotes.bas: smfGet888OptionQuote legacy endpoint", _
            "2026-11-09")
    Set smfUnsupportedRegistry = rows
End Function

Private Function smfUnsupportedRoutes() As Collection
    Dim rows As Collection
    Set rows = New Collection
    rows.Add Array("msn-options", "investing.money.msn.com", "/investments/equity-options/")
    rows.Add Array("msn-options", "investing.money.msn.com", "/investments/equity-options")
    rows.Add Array("optionsxpress-options", "www.optionsxpress.com", "/OXNetTools/Chains/index.aspx")
    rows.Add Array("optionsxpress-options", "optionsxpress.com", "/OXNetTools/Chains/index.aspx")
    rows.Add Array("yahoo-quote-csv", "download.finance.yahoo.com", "/d/quotes.csv")
    rows.Add Array("yahoo-yql", "query.yahooapis.com", "/v1/public/yql")
    rows.Add Array("yahoo-yql", "query.yahooapis.com", "/v1/public/yql/")
    rows.Add Array("google-option-chain", "www.google.com", "/finance/option_chain")
    rows.Add Array("google-option-chain", "finance.google.com", "/finance/option_chain")
    rows.Add Array("oic-legacy-options", "oic.ivolatility.com", "/oic_adv_options.j")
    Set smfUnsupportedRoutes = rows
End Function
' END GENERATED UNSUPPORTED SOURCES
Public Sub smfUnsupportedValue(ByVal pID As String, ByRef pValue As Variant)
    Dim rows As Collection, row As Variant
    Set rows = smfUnsupportedRegistry()
    For Each row In rows
        If StrComp(CStr(row(0)), pID, vbBinaryCompare) = 0 Then
            pValue = CVErr(xlErrNA)
            Exit Sub
        End If
    Next row
    Err.Raise 5, "smfUnsupportedValue", "Unknown unsupported-source ID"
End Sub

Public Sub smfUnsupportedURL(ByVal pURL As String, ByRef pID As String, ByRef pReason As String)
    Dim host As String, path As String, routes As Collection, rows As Collection
    Dim route As Variant, row As Variant
    pID = "": pReason = ""
    If Not smfUnsupportedURLParts(pURL, host, path) Then Exit Sub
    Set routes = smfUnsupportedRoutes()
    For Each route In routes
        If StrComp(host, CStr(route(1)), vbTextCompare) = 0 And _
                StrComp(path, CStr(route(2)), vbTextCompare) = 0 Then
            Set rows = smfUnsupportedRegistry()
            For Each row In rows
                If CStr(row(0)) = CStr(route(0)) Then
                    pID = CStr(row(0)): pReason = CStr(row(5))
                    Exit Sub
                End If
            Next row
        End If
    Next route
End Sub

Private Function smfUnsupportedURLParts(ByVal pURL As String, ByRef pHost As String, ByRef pPath As String) As Boolean
    Dim schemeAt As Long, authorityAt As Long, pathAt As Long, endAt As Long, pos As Long
    Dim authority As String, character As String
    schemeAt = InStr(1, pURL, "://", vbBinaryCompare)
    If schemeAt = 0 Then Exit Function
    If LCase$(Left$(pURL, schemeAt - 1)) <> "http" And _
            LCase$(Left$(pURL, schemeAt - 1)) <> "https" Then Exit Function
    authorityAt = schemeAt + 3
    endAt = Len(pURL) + 1
    For pos = authorityAt To Len(pURL)
        character = Mid$(pURL, pos, 1)
        If character = "/" Or character = "?" Or character = "#" Then endAt = pos: Exit For
    Next pos
    authority = Mid$(pURL, authorityAt, endAt - authorityAt)
    pos = InStrRev(authority, "@")
    If pos > 0 Then authority = Mid$(authority, pos + 1)
    If Left$(authority, 1) = "[" Then Exit Function
    pos = InStr(1, authority, ":", vbBinaryCompare)
    If pos > 0 Then authority = Left$(authority, pos - 1)
    If Right$(authority, 1) = "." Then authority = Left$(authority, Len(authority) - 1)
    If Len(authority) = 0 Then Exit Function
    pHost = LCase$(authority)
    pPath = "/"
    If endAt <= Len(pURL) Then
        If Mid$(pURL, endAt, 1) = "/" Then
            pathAt = endAt: endAt = Len(pURL) + 1
            For pos = pathAt To Len(pURL)
                character = Mid$(pURL, pos, 1)
                If character = "?" Or character = "#" Then endAt = pos: Exit For
            Next pos
            pPath = Mid$(pURL, pathAt, endAt - pathAt)
        End If
    End If
    smfUnsupportedURLParts = True
End Function

Private Function smfLiteralUnsupportedReason(ByVal pExpression As String, ByRef pReason As String) As Boolean
    Dim expression As String, literal As String, pos As Long, character As String, quoted As Boolean
    pReason = ""
    expression = Trim$(pExpression)
    If Left$(expression, 1) = "=" Then
        expression = Trim$(Mid$(expression, 2))
        If Left$(expression, 1) <> Chr$(34) Then Exit Function
    End If
    If Left$(expression, 1) = Chr$(34) Then
        If Len(expression) < 2 Or Right$(expression, 1) <> Chr$(34) Then Exit Function
        pos = 2
        Do While pos < Len(expression)
            character = Mid$(expression, pos, 1)
            If character = Chr$(34) Then
                If pos + 1 >= Len(expression) Then Exit Function
                If Mid$(expression, pos + 1, 1) <> Chr$(34) Then Exit Function
                pos = pos + 1
            End If
            literal = literal & character
            pos = pos + 1
        Loop
    Else
        literal = expression
    End If
    Select Case LCase$(literal)
        Case "obsolete": pReason = "The effective loaded SMF definition explicitly marks this field obsolete."
        Case "obsolete -- web site no longer exists": pReason = "The effective loaded SMF definition records that the source web site no longer exists."
        Case Else: Exit Function
    End Select
    smfLiteralUnsupportedReason = True
End Function

Public Sub smfUnsupportedRows(ByRef pRows As Variant)
    Dim records As Collection, record As Variant, i As Long, j As Long, reason As String
    Set records = smfUnsupportedRegistry()
    If iInit <> 0 Then
        For i = 1 To kElements
            If Len(aParms(i)) > 0 Then
                Dim parts As Variant
                parts = Split(aParms(i), ";")
                If UBound(parts) >= 2 Then
                    If smfLiteralUnsupportedReason(CStr(parts(2)), reason) Then
                        records.Add Array("field-" & CStr(i), "RCHGetElementNumber field " & CStr(i) & ": " & CStr(parts(1)), _
                            "Current loaded numbered definition", CStr(parts(0)), "Not Supported", reason, _
                            "Supply a replacement definition or use another supported source.", aFiles(i), "2026-11-09")
                    End If
                End If
            End If
        Next i
    Else
        records.Add Array("definition-scope", "RCHGetElementNumber", "Definition inventory", "Local/Internet as configured", _
            "Not Loaded", "This table reads memory only. Numbered-field rows appear after SMF loads definitions.", _
            "The shipped-default snapshot is mac/NotSupported.csv.", "No definition initialization was requested.", "2026-11-09")
    End If
    ReDim pRows(1 To records.Count + 1, 1 To 9)
    record = Array("ID", "Function or field", "Mode", "Provider", "Status", "Reason", "Replacement", "Evidence", "Reviewed")
    For j = 0 To 8: pRows(1, j + 1) = record(j): Next j
    i = 1
    For Each record In records
        i = i + 1
        For j = 0 To 8: pRows(i, j + 1) = record(j): Next j
    Next record
End Sub

Public Sub smfShowSourceStatus()
    Dim rows As Variant, book As Workbook, sheet As Worksheet, table As ListObject
    smfUnsupportedRows rows
    Set book = Application.Workbooks.Add(xlWBATWorksheet)
    Set sheet = book.Worksheets(1)
    sheet.Name = "SMFSourceStatus"
    With sheet
        .Range("A1").Resize(UBound(rows, 1), UBound(rows, 2)).NumberFormat = "@"
        .Range("A1").Resize(UBound(rows, 1), UBound(rows, 2)).Value2 = rows
        .Rows(1).Font.Bold = True
        Set table = .ListObjects.Add(xlSrcRange, .Range("A1:I" & CStr(UBound(rows, 1))), , xlYes)
        table.Name = "SMFSourceStatusTable"
        table.ShowAutoFilter = True
        .Columns("A").ColumnWidth = 28
        .Columns("B").ColumnWidth = 42
        .Columns("C:D").ColumnWidth = 28
        .Columns("E").ColumnWidth = 18
        .Columns("F:H").ColumnWidth = 58
        .Columns("I").ColumnWidth = 14
        .UsedRange.WrapText = True
        .UsedRange.Rows.AutoFit
    End With
End Sub

Private Function smfUnsupportedFieldSelfChecks() As Long
    Dim oldParm As String, oldFile As String, oldInit As Integer, oldError As Variant
    Dim value As Variant, rows As Variant, selectors As Variant, expected As Variant, literal As Variant
    Dim callsBefore As Double, requestsBefore As Double, callsAfter As Double, requestsAfter As Double
    Dim count As Long, i As Long, found As Boolean, errorNumber As Long, errorText As String
    oldParm = aParms(19999): oldFile = aFiles(19999): oldInit = iInit
    If IsObject(vError) Then Set oldError = vError Else oldError = vError
    On Error GoTo Failed
    iInit = 1
    aParms(19999) = "Fixture;Retired fixture;=""Obsolete"""
    aFiles(19999) = "unsupported-field-fixture.txt"
    SMFMacRequestCounts callsBefore, requestsBefore
    value = RCHGetElementNumber("AAPL", 19999, "CUSTOM_FALLBACK", "NO_WORKBOOK_LOOKUP")
    smfUnsupportedAssertNA value
    count = count + 1
    selectors = Array("DEFINITION", "SOURCE", "ELEMENT", "ELEMENTFILE", "WEB PAGE", "P-URL")
    expected = Array(aParms(19999), "Fixture", "Retired fixture", aFiles(19999), "=""Obsolete""", "=""Obsolete""")
    For i = 0 To UBound(selectors)
        value = RCHGetElementNumber(CStr(selectors(i)), 19999)
        If VarType(value) <> vbString Then Err.Raise 5, , "Retired-field metadata lost its String type"
        If CStr(value) <> CStr(expected(i)) Then Err.Raise 5, , "Retired-field metadata changed"
        count = count + 1
    Next i
    aParms(19999) = "Fixture;Retired fixture;=""Obsolete -- Web site no longer exists"""
    value = RCHGetElementNumber("AAPL", 19999)
    smfUnsupportedAssertNA value
    count = count + 1
    For Each literal In Array("Obsolete -- No data found", "Obsolete -- No longer available?", _
            "Obsolete -- Now only available to premium subscribers")
        aParms(19999) = "Fixture;Retired fixture;=" & Chr$(34) & CStr(literal) & Chr$(34)
        value = RCHGetElementNumber("AAPL", 19999)
        If VarType(value) <> vbString Then Err.Raise 5, , "An uncertain/premium definition was disabled"
        If CStr(value) <> CStr(literal) Then Err.Raise 5, , "An uncertain/premium definition changed"
        count = count + 1
    Next literal
    aParms(19999) = "Fixture;Conditional fixture;=IF(1=1,""Obsolete"",42)"
    value = RCHGetElementNumber("AAPL", 19999)
    If VarType(value) <> vbString Then Err.Raise 5, , "A conditional definition was disabled"
    If CStr(value) <> "Obsolete" Then Err.Raise 5, , "A conditional definition changed"
    count = count + 1
    aParms(19999) = "Fixture;Retired fixture;=""Obsolete"""
    smfUnsupportedRows rows
    found = False
    For i = 2 To UBound(rows, 1)
        If CStr(rows(i, 1)) = "field-19999" Then
            found = True
            If CStr(rows(i, 5)) <> "Not Supported" Then Err.Raise 5, , "Retired field missing its status"
            If CStr(rows(i, 8)) <> aFiles(19999) Then Err.Raise 5, , "Retired-field file metadata changed"
        End If
    Next i
    If Not found Then Err.Raise 5, , "The table omitted an effective retired definition"
    count = count + 1
    aParms(19999) = "Fixture;Replacement fixture;=1+2"
    value = RCHGetElementNumber("AAPL", 19999)
    If Not IsNumeric(value) Then Err.Raise 5, , "A working replacement definition was disabled"
    If CDbl(value) <> 3 Then Err.Raise 5, , "A working replacement definition returned the wrong value"
    count = count + 1
    smfUnsupportedRows rows
    For i = 2 To UBound(rows, 1)
        If CStr(rows(i, 1)) = "field-19999" Then Err.Raise 5, , "The table retained a replaced field"
    Next i
    count = count + 1
    SMFMacRequestCounts callsAfter, requestsAfter
    If callsAfter <> callsBefore Or requestsAfter <> requestsBefore Then Err.Raise 5, , "A retired-field check reached the downloader"
    count = count + 1
    smfUnsupportedFieldSelfChecks = count
Restore:
    aParms(19999) = oldParm: aFiles(19999) = oldFile: iInit = oldInit
    If IsObject(oldError) Then Set vError = oldError Else vError = oldError
    If errorNumber <> 0 Then Err.Raise errorNumber, "smfUnsupportedFieldSelfChecks", errorText
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Resume Restore
End Function

Private Sub smfUnsupportedAssertNA(ByVal pValue As Variant)
    If Not IsError(pValue) Then Err.Raise 5, , "A retired field did not return an Excel error"
    If CStr(pValue) <> CStr(CVErr(xlErrNA)) Then Err.Raise 5, , "A retired field returned a different Excel error"
End Sub

#End If

Attribute VB_Name = "modMacHTML"
Option Explicit
Option Compare Binary

' HTMLDocument.body.innerText replacement for the text fragments used by
' smfGetTagContent. This module uses only VBA string operations.
' It does not execute JavaScript, apply CSS, or repair malformed HTML as a browser does.
' Table cells use tabs; rows/blocks use CRLF; paragraphs use two CRLFs.
' HTML4 named entities plus common HTML5 punctuation names are supported.
Private Type smfMacHTMLState
    textBuffer As String
    textLength As Long
    pendingSpace As Boolean
    pendingLines As Long
    preDepth As Long
    preBreak As String
    cellsInRow As Long
End Type

Public Function smfMacStripHTML(ByVal pString As String) As String
    Dim state As smfMacHTMLState
    Dim pos As Long, tagEnd As Long, commentEnd As Long, closeAt As Long
    Dim tagName As String, skippedTag As String, decoded As String
    Dim isClosing As Boolean, isSelfClosing As Boolean

    state.textBuffer = Space$(Len(pString) + 2)
    pos = 1
    Do While pos <= Len(pString)
        If Len(skippedTag) > 0 Then
            ' Raw script/style text may contain comparison signs and tag-looking strings.
            closeAt = InStr(pos, pString, "</" & skippedTag, vbTextCompare)
            Do While closeAt > 0
                If smfMacHTMLReadTag(pString, closeAt, tagEnd, tagName, isClosing, isSelfClosing) Then
                    If isClosing And tagName = skippedTag Then Exit Do
                End If
                closeAt = InStr(closeAt + 2, pString, "</" & skippedTag, vbTextCompare)
            Loop
            If closeAt = 0 Then Exit Do
            skippedTag = ""
            pos = tagEnd + 1
            GoTo NextCharacter
        End If

        If Mid$(pString, pos, 4) = "<!--" Then
            commentEnd = InStr(pos + 4, pString, "-->", vbBinaryCompare)
            If commentEnd = 0 Then Exit Do
            pos = commentEnd + 3
            GoTo NextCharacter
        End If

        If Mid$(pString, pos, 1) = "<" Then
            If smfMacHTMLReadTag(pString, pos, tagEnd, tagName, isClosing, isSelfClosing) Then
                If Not isClosing And Not isSelfClosing Then
                    Select Case tagName
                        Case "script", "style", "head", "title", "noscript"
                            skippedTag = tagName
                    End Select
                End If
                If Len(skippedTag) = 0 Then
                    smfMacHTMLHandleTag state, tagName, isClosing, isSelfClosing
                End If
                pos = tagEnd + 1
                GoTo NextCharacter
            End If
        End If

        If Mid$(pString, pos, 1) = "&" Then
            If smfMacHTMLReadEntity(pString, pos, tagEnd, decoded) Then
                ' Decode once: &amp;lt; is the literal text &lt;, not a new tag.
                smfMacHTMLAppendText state, decoded
                pos = tagEnd + 1
                GoTo NextCharacter
            End If
        End If

        smfMacHTMLAppendText state, Mid$(pString, pos, 1)
        pos = pos + 1
NextCharacter:
    Loop
    smfMacStripHTML = Left$(state.textBuffer, state.textLength)
End Function

Private Function smfMacHTMLReadTag(ByVal pHTML As String, ByVal pAt As Long, _
        ByRef pEnd As Long, ByRef pName As String, ByRef pClosing As Boolean, _
        ByRef pSelfClosing As Boolean) As Boolean
    Dim pos As Long, nameStart As Long, char As String, quote As String
    pClosing = False
    pSelfClosing = False
    pName = ""
    pos = pAt + 1
    char = Mid$(pHTML, pos, 1)
    If char = "/" Then
        pClosing = True
        pos = pos + 1
        char = Mid$(pHTML, pos, 1)
    End If
    If char = "!" Or char = "?" Then
        pName = "!"
    Else
        ' A space/digit after '<' is text, as in "debt < 5" or "yield <3%".
        If Not smfMacHTMLIsLetter(char) Then Exit Function
        nameStart = pos
        Do While pos <= Len(pHTML)
            char = Mid$(pHTML, pos, 1)
            If Not smfMacHTMLIsNameChar(char) Then Exit Do
            pos = pos + 1
        Loop
        pName = LCase$(Mid$(pHTML, nameStart, pos - nameStart))
        char = Mid$(pHTML, pos, 1)
        If char <> ">" And char <> "/" And Not smfMacHTMLIsSpace(char) Then Exit Function
    End If

    Do While pos <= Len(pHTML)
        char = Mid$(pHTML, pos, 1)
        If Len(quote) > 0 Then
            If char = quote Then quote = ""
        ElseIf char = Chr$(34) Or char = "'" Then
            quote = char
        ElseIf char = ">" Then
            pEnd = pos
            pos = pos - 1
            Do While pos > pAt
                If Not smfMacHTMLIsSpace(Mid$(pHTML, pos, 1)) Then Exit Do
                pos = pos - 1
            Loop
            pSelfClosing = (Mid$(pHTML, pos, 1) = "/")
            smfMacHTMLReadTag = True
            Exit Function
        End If
        pos = pos + 1
    Loop
End Function

Private Sub smfMacHTMLHandleTag(ByRef pState As smfMacHTMLState, ByVal pName As String, _
        ByVal pClosing As Boolean, ByVal pSelfClosing As Boolean)
    Select Case pName
        Case "br"
            pState.pendingSpace = False
            smfMacHTMLFlushPending pState
            smfMacHTMLAppendRaw pState, vbCrLf
            pState.preBreak = ""
        Case "td", "th"
            pState.pendingSpace = False
            If Not pClosing Then
                If pState.cellsInRow > 0 Then
                    smfMacHTMLFlushPending pState
                    smfMacHTMLAppendRaw pState, vbTab
                End If
                pState.cellsInRow = pState.cellsInRow + 1
            End If
        Case "tr", "table"
            pState.cellsInRow = 0
            smfMacHTMLSetBreak pState, 1
        Case "pre"
            smfMacHTMLSetBreak pState, 1
            If pClosing Then
                If pState.preDepth > 0 Then pState.preDepth = pState.preDepth - 1
            ElseIf Not pSelfClosing Then
                pState.preDepth = pState.preDepth + 1
            End If
        Case "p"
            smfMacHTMLSetBreak pState, 2
        Case "address", "article", "aside", "blockquote", "caption", "center", _
             "dd", "details", "dialog", "dir", "div", "dl", "dt", "fieldset", _
             "figcaption", "figure", "footer", "form", "h1", "h2", "h3", "h4", _
             "h5", "h6", "header", "hgroup", "hr", "li", "main", "menu", "nav", _
             "ol", "section", "summary", "tbody", "tfoot", "thead", "ul"
            smfMacHTMLSetBreak pState, 1
    End Select
End Sub

Private Sub smfMacHTMLSetBreak(ByRef pState As smfMacHTMLState, ByVal pLines As Long)
    If pLines > pState.pendingLines Then pState.pendingLines = pLines
    pState.pendingSpace = False
    pState.preBreak = ""
End Sub

Private Sub smfMacHTMLAppendText(ByRef pState As smfMacHTMLState, ByVal pText As String)
    Dim pos As Long, char As String
    For pos = 1 To Len(pText)
        char = Mid$(pText, pos, 1)
        If pState.preDepth > 0 Then
            smfMacHTMLFlushPending pState
            Select Case char
                Case Chr$(13), Chr$(10)
                    ' Normalize CR, LF, CRLF and LFCR to the existing vbCrLf output.
                    ' Clear the state after a pair so CRLFCRLF keeps two breaks.
                    If Len(pState.preBreak) > 0 And pState.preBreak <> char Then
                        pState.preBreak = ""
                    Else
                        smfMacHTMLAppendRaw pState, vbCrLf
                        pState.preBreak = char
                    End If
                Case Else
                    smfMacHTMLAppendRaw pState, char
                    pState.preBreak = ""
            End Select
        ElseIf smfMacHTMLIsSpace(char) Then
            pState.pendingSpace = True
        Else
            smfMacHTMLFlushPending pState
            smfMacHTMLAppendRaw pState, char
        End If
    Next pos
End Sub

Private Sub smfMacHTMLFlushPending(ByRef pState As smfMacHTMLState)
    Dim linesAlready As Long, pos As Long, lastChar As String
    If pState.textLength > 0 Then
        If pState.pendingLines > 0 Then
            pos = pState.textLength - 1
            Do While pos > 0
                If Mid$(pState.textBuffer, pos, 2) <> vbCrLf Then Exit Do
                linesAlready = linesAlready + 1
                pos = pos - 2
            Loop
            Do While linesAlready < pState.pendingLines
                smfMacHTMLAppendRaw pState, vbCrLf
                linesAlready = linesAlready + 1
            Loop
        ElseIf pState.pendingSpace Then
            lastChar = Mid$(pState.textBuffer, pState.textLength, 1)
            If lastChar <> vbLf And lastChar <> vbTab Then smfMacHTMLAppendRaw pState, " "
        End If
    End If
    pState.pendingLines = 0
    pState.pendingSpace = False
End Sub

Private Sub smfMacHTMLAppendRaw(ByRef pState As smfMacHTMLState, ByVal pText As String)
    Dim required As Long, capacity As Long
    required = pState.textLength + Len(pText)
    capacity = Len(pState.textBuffer)
    If required > capacity Then
        If capacity < 128 Then capacity = 128
        Do While capacity < required
            capacity = capacity * 2
        Loop
        pState.textBuffer = pState.textBuffer & Space$(capacity - Len(pState.textBuffer))
    End If
    If Len(pText) > 0 Then Mid$(pState.textBuffer, pState.textLength + 1, Len(pText)) = pText
    pState.textLength = required
End Sub

Private Function smfMacHTMLIsSpace(ByVal pChar As String) As Boolean
    Select Case pChar
        Case " ", vbTab, vbCr, vbLf, Chr$(12)
            smfMacHTMLIsSpace = True
    End Select
End Function

Private Function smfMacHTMLIsLetter(ByVal pChar As String) As Boolean
    If Len(pChar) <> 1 Then Exit Function
    smfMacHTMLIsLetter = (pChar >= "a" And pChar <= "z") Or (pChar >= "A" And pChar <= "Z")
End Function

Private Function smfMacHTMLIsNameChar(ByVal pChar As String) As Boolean
    If Len(pChar) <> 1 Then Exit Function
    smfMacHTMLIsNameChar = smfMacHTMLIsLetter(pChar) Or _
        (pChar >= "0" And pChar <= "9") Or pChar = "-" Or pChar = ":" Or pChar = "_"
End Function

Private Function smfMacHTMLReadEntity(ByVal pHTML As String, ByVal pAt As Long, _
        ByRef pEnd As Long, ByRef pValue As String) As Boolean
    Dim pos As Long, start As Long, radix As Long, digit As Long, codePoint As Long
    Dim char As String, entityName As String, outOfRange As Boolean
    pos = pAt + 1
    If Mid$(pHTML, pos, 1) = "#" Then
        pos = pos + 1
        radix = 10
        If LCase$(Mid$(pHTML, pos, 1)) = "x" Then
            radix = 16
            pos = pos + 1
        End If
        start = pos
        Do While pos <= Len(pHTML)
            digit = smfMacHTMLDigit(Mid$(pHTML, pos, 1), radix)
            If digit < 0 Then Exit Do
            If Not outOfRange Then
                If codePoint > (1114111 - digit) \ radix Then
                    outOfRange = True
                Else
                    codePoint = codePoint * radix + digit
                End If
            End If
            pos = pos + 1
        Loop
        If pos = start Then Exit Function
        If outOfRange Then codePoint = 65533
        pValue = smfMacHTMLUnicode(codePoint)
    Else
        start = pos
        Do While pos <= Len(pHTML)
            char = Mid$(pHTML, pos, 1)
            If Not smfMacHTMLIsLetter(char) And Not (char >= "0" And char <= "9") Then Exit Do
            pos = pos + 1
            If pos - start > 32 Then Exit Function
        Loop
        entityName = Mid$(pHTML, start, pos - start)
        If Len(entityName) = 0 Then Exit Function
        If Mid$(pHTML, pos, 1) <> ";" Then
            ' Support common legacy entities without a semicolon in text nodes.
            Select Case entityName
                Case "amp", "AMP", "lt", "LT", "gt", "GT", "quot", "QUOT", "nbsp", "copy", "COPY", "reg", "REG"
                Case Else
                    Exit Function
            End Select
        End If
        If Not smfMacHTMLNamedEntity(entityName, pValue) Then Exit Function
    End If
    If Mid$(pHTML, pos, 1) = ";" Then pos = pos + 1
    pEnd = pos - 1
    smfMacHTMLReadEntity = True
End Function

Private Function smfMacHTMLDigit(ByVal pChar As String, ByVal pRadix As Long) As Long
    Dim code As Long
    smfMacHTMLDigit = -1
    If Len(pChar) <> 1 Then Exit Function
    code = AscW(pChar)
    If code >= 48 And code <= 57 Then
        smfMacHTMLDigit = code - 48
    ElseIf pRadix = 16 Then
        If code >= 65 And code <= 70 Then smfMacHTMLDigit = code - 55
        If code >= 97 And code <= 102 Then smfMacHTMLDigit = code - 87
    End If
End Function

Private Function smfMacHTMLNamedEntity(ByVal pName As String, ByRef pValue As String) As Boolean
    Static names As Variant, codes As Variant
    Dim i As Long
    If IsEmpty(names) Then
        names = Split("AElig|AMP|Aacute|Acirc|Agrave|Alpha|Aring|Atilde|Auml|Beta|COPY|Ccedil|Chi|Dagger|Delta|ETH|Eacute|Ecirc|Egrave|Epsilon|Eta|Euml|GT|Gamma|Iacute|" & _
        "Icirc|Igrave|Iota|Iuml|Kappa|LT|Lambda|Mu|NewLine|Ntilde|Nu|OElig|Oacute|Ocirc|Ograve|Omega|Omicron|Oslash|Otilde|Ouml|Phi|Pi|Prime|Psi|QUOT|" & _
        "REG|Rho|Scaron|Sigma|THORN|Tab|Tau|Theta|Uacute|Ucirc|Ugrave|Upsilon|Uuml|Xi|Yacute|Yuml|ZeroWidthSpace|Zeta|aacute|acirc|acute|aelig|agrave|alefsym|alpha|" & _
        "amp|and|ang|apos|aring|asymp|atilde|auml|bdquo|beta|brvbar|bsol|bull|cap|ccedil|cedil|cent|chi|circ|clubs|colon|comma|commat|cong|copy|" & _
        "crarr|cup|curren|dArr|dagger|darr|deg|delta|diams|divide|dollar|eacute|ecirc|egrave|empty|emsp|ensp|epsilon|equals|equiv|eta|eth|euml|euro|exist|" & _
        "fnof|forall|frac12|frac14|frac34|frasl|gamma|ge|gt|hArr|harr|hearts|hellip|iacute|icirc|iexcl|igrave|image|infin|int|iota|iquest|isin|iuml|kappa|" & _
        "lArr|lambda|lang|laquo|larr|lceil|ldquo|le|lfloor|lowast|loz|lpar|lrm|lsaquo|lsquo|lt|macr|mdash|micro|middot|minus|mu|nabla|nbsp|ndash|" & _
        "ne|ni|not|notin|nsub|ntilde|nu|num|oacute|ocirc|oelig|ograve|oline|omega|omicron|oplus|or|ordf|ordm|oslash|otilde|otimes|ouml|para|part|" & _
        "percnt|period|permil|perp|phi|pi|piv|plus|plusmn|pound|prime|prod|prop|psi|quot|rArr|radic|rang|raquo|rarr|rceil|rdquo|real|reg|rfloor|" & _
        "rho|rlm|rpar|rsaquo|rsquo|sbquo|scaron|sdot|sect|semi|shy|sigma|sigmaf|sim|sol|spades|sub|sube|sum|sup|sup1|sup2|sup3|supe|szlig|" & _
        "tau|there4|theta|thetasym|thinsp|thorn|tilde|times|trade|uArr|uacute|uarr|ucirc|ugrave|uml|upsih|upsilon|uuml|weierp|xi|yacute|yen|yuml|zeta|zwj|" & _
        "zwnj", "|")
        codes = Split("198|38|193|194|192|913|197|195|196|914|169|199|935|8225|916|208|201|202|200|917|919|203|62|915|205|" & _
        "206|204|921|207|922|60|923|924|10|209|925|338|211|212|210|937|927|216|213|214|934|928|8243|936|34|" & _
        "174|929|352|931|222|9|932|920|218|219|217|933|220|926|221|376|8203|918|225|226|180|230|224|8501|945|" & _
        "38|8743|8736|39|229|8776|227|228|8222|946|166|92|8226|8745|231|184|162|967|710|9827|58|44|64|8773|169|" & _
        "8629|8746|164|8659|8224|8595|176|948|9830|247|36|233|234|232|8709|8195|8194|949|61|8801|951|240|235|8364|8707|" & _
        "402|8704|189|188|190|8260|947|8805|62|8660|8596|9829|8230|237|238|161|236|8465|8734|8747|953|191|8712|239|954|" & _
        "8656|955|9001|171|8592|8968|8220|8804|8970|8727|9674|40|8206|8249|8216|60|175|8212|181|183|8722|956|8711|160|8211|" & _
        "8800|8715|172|8713|8836|241|957|35|243|244|339|242|8254|969|959|8853|8744|170|186|248|245|8855|246|182|8706|" & _
        "37|46|8240|8869|966|960|982|43|177|163|8242|8719|8733|968|34|8658|8730|9002|187|8594|8969|8221|8476|174|8971|" & _
        "961|8207|41|8250|8217|8218|353|8901|167|59|173|963|962|8764|47|9824|8834|8838|8721|8835|185|178|179|8839|223|" & _
        "964|8756|952|977|8201|254|732|215|8482|8657|250|8593|251|249|168|978|965|252|8472|958|253|165|255|950|8205|" & _
        "8204", "|")
    End If
    For i = LBound(names) To UBound(names)
        If StrComp(pName, names(i), vbBinaryCompare) = 0 Then
            pValue = smfMacHTMLUnicode(CLng(codes(i)))
            smfMacHTMLNamedEntity = True
            Exit Function
        End If
    Next i
End Function

Private Function smfMacHTMLUnicode(ByVal pCodePoint As Long) As String
    Dim highSurrogate As Long, lowSurrogate As Long
    ' HTML numeric references map these legacy Windows-1252 values to Unicode.
    Select Case pCodePoint
        Case 128: pCodePoint = 8364
        Case 130: pCodePoint = 8218
        Case 131: pCodePoint = 402
        Case 132: pCodePoint = 8222
        Case 133: pCodePoint = 8230
        Case 134: pCodePoint = 8224
        Case 135: pCodePoint = 8225
        Case 136: pCodePoint = 710
        Case 137: pCodePoint = 8240
        Case 138: pCodePoint = 352
        Case 139: pCodePoint = 8249
        Case 140: pCodePoint = 338
        Case 142: pCodePoint = 381
        Case 145: pCodePoint = 8216
        Case 146: pCodePoint = 8217
        Case 147: pCodePoint = 8220
        Case 148: pCodePoint = 8221
        Case 149: pCodePoint = 8226
        Case 150: pCodePoint = 8211
        Case 151: pCodePoint = 8212
        Case 152: pCodePoint = 732
        Case 153: pCodePoint = 8482
        Case 154: pCodePoint = 353
        Case 155: pCodePoint = 8250
        Case 156: pCodePoint = 339
        Case 158: pCodePoint = 382
        Case 159: pCodePoint = 376
    End Select
    If pCodePoint = 0 Or pCodePoint > 1114111 Or _
            (pCodePoint >= 55296 And pCodePoint <= 57343) Then pCodePoint = 65533
    If pCodePoint > 65535 Then
        pCodePoint = pCodePoint - 65536
        highSurrogate = 55296 + (pCodePoint \ 1024)
        lowSurrogate = 56320 + (pCodePoint And 1023)
        smfMacHTMLUnicode = ChrW(highSurrogate - 65536) & ChrW(lowSurrogate - 65536)
    Else
        If pCodePoint > 32767 Then pCodePoint = pCodePoint - 65536
        smfMacHTMLUnicode = ChrW(pCodePoint)
    End If
End Function

' Run from Excel's VBA editor or Application.Run("smfMacHTMLSelfTest").
' Returns the number of checks; raises an error with the failing case name.
Public Function smfMacHTMLSelfTest() As Long
    Dim passed As Long
    smfMacHTMLAssert "empty", "", "", passed
    smfMacHTMLAssert "plain financial text", "Revenue $1,234.50 (5%)", "Revenue $1,234.50 (5%)", passed
    smfMacHTMLAssert "comparisons", "Revenue > 5 and debt < 3", "Revenue > 5 and debt < 3", passed
    smfMacHTMLAssert "spaced angle signs", "A < B > C", "A < B > C", passed
    smfMacHTMLAssert "compact numeric comparison", "yield <3%", "yield <3%", passed
    smfMacHTMLAssert "unclosed tag stays text", "price <span", "price <span", passed
    smfMacHTMLAssert "inline financial number", "$<b>1,</b><i>234.50</i>", "$1,234.50", passed
    smfMacHTMLAssert "collapse source whitespace", "  Revenue" & vbCrLf & vbTab & "growth  ", "Revenue growth", passed
    smfMacHTMLAssert "nested tags", "<div>A<span>B</span></div><div>C</div>", "AB" & vbCrLf & "C", passed
    smfMacHTMLAssert "paragraphs", "<p>A</p><p>B</p>", "A" & vbCrLf & vbCrLf & "B", passed
    smfMacHTMLAssert "line breaks", "A<br>B<br/>C", "A" & vbCrLf & "B" & vbCrLf & "C", passed
    smfMacHTMLAssert "repeated line breaks", "A<br><br>B", "A" & vbCrLf & vbCrLf & "B", passed
    smfMacHTMLAssert "quoted greater sign", "<span title='x > y'>7</span>", "7", passed
    smfMacHTMLAssert "quoted markup", "<span title=""<b> > x"">7</span>", "7", passed
    smfMacHTMLAssert "comments", "A<!-- ignored <td> -->B", "AB", passed
    smfMacHTMLAssert "unclosed comment", "A<!-- ignored", "A", passed
    smfMacHTMLAssert "script raw content", "A<script>if (a < b) {x='<td>9</td>';}</script>B", "AB", passed
    smfMacHTMLAssert "style raw content", "<style>.a > b {content:'<tr>';}</style>42", "42", passed
    smfMacHTMLAssert "case insensitive tags", "<ScRiPt>999</sCrIpT><B>42</b>", "42", passed
    smfMacHTMLAssert "script closing name boundary", "<script>'</scripture>'</script>42", "42", passed
    smfMacHTMLAssert "head removed", "<!DOCTYPE html><html><head><title>ignore</title></head><body>42</body></html>", "42", passed
    smfMacHTMLAssert "table", "<table><tr><th>Metric</th><th>Value</th></tr><tr><td>EPS</td><td>$1.25</td></tr></table>", "Metric" & vbTab & "Value" & vbCrLf & "EPS" & vbTab & "$1.25", passed
    smfMacHTMLAssert "omitted cell closing tags", "<table><tr><td>A<td>B<tr><td>C<td>D</table>", "A" & vbTab & "B" & vbCrLf & "C" & vbTab & "D", passed
    smfMacHTMLAssert "empty first cell", "<tr><td></td><td>10</td></tr>", vbTab & "10", passed
    smfMacHTMLAssert "escaped comparisons", "&lt;5 &amp;&amp; &gt;3", "<5 && >3", passed
    smfMacHTMLAssert "decode once", "&amp;lt;b&amp;gt;", "&lt;b&gt;", passed
    smfMacHTMLAssert "quotes and apostrophes", "&quot;A&apos;B&quot;", Chr$(34) & "A'B" & Chr$(34), passed
    smfMacHTMLAssert "nonbreaking space", "A&nbsp;B", "A" & ChrW(160) & "B", passed
    smfMacHTMLAssert "financial symbols", "&pound;1 &euro;2 &yen;3 &minus;4", ChrW(163) & "1 " & ChrW(8364) & "2 " & ChrW(165) & "3 " & ChrW(8722) & "4", passed
    smfMacHTMLAssert "named financial punctuation", "&dollar;1&comma;234&period;50&percnt;", "$1,234.50%", passed
    smfMacHTMLAssert "decimal numeric entities", "&#36;1,234&#46;50", "$1,234.50", passed
    smfMacHTMLAssert "hex numeric entities", "&#x24;1&#X2E;5", "$1.5", passed
    smfMacHTMLAssert "legacy numeric dash", "&#151;", ChrW(8212), passed
    smfMacHTMLAssert "Greek names", "&Alpha;&alpha;&beta;", ChrW(913) & ChrW(945) & ChrW(946), passed
    smfMacHTMLAssert "case sensitive names", "&Aacute;&aacute;", ChrW(193) & ChrW(225), passed
    smfMacHTMLAssert "unknown entity", "A&notARealEntity;B", "A&notARealEntity;B", passed
    smfMacHTMLAssert "literal ampersand", "AT&T", "AT&T", passed
    smfMacHTMLAssert "legacy omitted semicolon", "A &amp B &#36 C", "A & B $ C", passed
    smfMacHTMLAssert "invalid numeric entity", "&#xZZ;", "&#xZZ;", passed
    smfMacHTMLAssert "out of range entity", "&#999999999999999999999;", ChrW(-3), passed
    smfMacHTMLAssert "null entity", "&#0;", ChrW(-3), passed
    smfMacHTMLAssert "surrogate entity", "&#xD800;", ChrW(-3), passed
    smfMacHTMLAssert "supplementary character", "&#x1F4B0;", ChrW(-10179) & ChrW(-9040), passed
    smfMacHTMLAssert "preformatted text", "<pre>  A" & vbLf & vbTab & "B  </pre>", "  A" & vbCrLf & vbTab & "B  ", passed
    smfMacHTMLAssert "preformatted CRLF", "<pre>A" & vbCrLf & "B</pre>", "A" & vbCrLf & "B", passed
    smfMacHTMLAssert "preformatted explicit CR", "<pre>A" & Chr$(13) & "B</pre>", "A" & vbCrLf & "B", passed
    smfMacHTMLAssert "preformatted explicit LF", "<pre>A" & Chr$(10) & "B</pre>", "A" & vbCrLf & "B", passed
    smfMacHTMLAssert "preformatted explicit CRLF", "<pre>A" & Chr$(13) & Chr$(10) & "B</pre>", "A" & vbCrLf & "B", passed
    smfMacHTMLAssert "preformatted explicit LFCR", "<pre>A" & Chr$(10) & Chr$(13) & "B</pre>", "A" & vbCrLf & "B", passed
    smfMacHTMLAssert "preformatted repeated CR", "<pre>A" & String$(2, Chr$(13)) & "B</pre>", "A" & vbCrLf & vbCrLf & "B", passed
    smfMacHTMLAssert "preformatted repeated LF", "<pre>A" & String$(2, Chr$(10)) & "B</pre>", "A" & vbCrLf & vbCrLf & "B", passed
    smfMacHTMLAssert "preformatted adjacent pairs", "<pre>A" & Chr$(13) & Chr$(10) & Chr$(10) & Chr$(13) & "B</pre>", "A" & vbCrLf & vbCrLf & "B", passed
    smfMacHTMLAssert "preformatted numeric newline entities", "<pre>A&#13;&#10;B</pre>", "A" & vbCrLf & "B", passed
    smfMacHTMLAssert "buffer growth", "<pre>" & String$(200, vbCr) & "</pre>", Replace(String$(200, vbCr), vbCr, vbCrLf), passed
    smfMacHTMLAssert "large fragment", "<div>" & String$(5000, "1") & "</div>", String$(5000, "1"), passed
    smfMacHTMLSelfTest = passed
End Function

Private Sub smfMacHTMLAssert(ByVal pName As String, ByVal pHTML As String, _
        ByVal pExpected As String, ByRef pPassed As Long)
    Dim actual As String
    actual = smfMacStripHTML(pHTML)
    If StrComp(actual, pExpected, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 2150, "smfMacHTMLSelfTest", _
            pName & ": expected [" & smfMacHTMLVisible(pExpected) & "] but got [" & smfMacHTMLVisible(actual) & "]"
    End If
    pPassed = pPassed + 1
End Sub

Private Function smfMacHTMLVisible(ByVal pText As String) As String
    smfMacHTMLVisible = Replace(Replace(Replace(pText, vbCr, "<CR>"), vbLf, "<LF>"), vbTab, "<TAB>")
End Function

Attribute VB_Name = "modJSONExtract"
Option Explicit

'-----------------------------------------------------------------------------------
' 08/08/23 - ScriptEngine/JSON extract code copied from
'            "https://stackoverflow.com/questions/6627652/parsing-json-in-excel-vba"
'-----------------------------------------------------------------------------------
' Windows x32 uses Script Control; Mac uses the bundled JsonConverter.
'-----------------------------------------------------------------------------------

#If Mac = 0 Then
' Late binding avoids a missing Script Control reference when imported on Mac.
Private ScriptEngine As Object
#End If

Public Sub InitScriptEngine()
#If Mac = 0 Then
    Set ScriptEngine = CreateObject("MSScriptControl.ScriptControl")
    ScriptEngine.Language = "JScript"
    ScriptEngine.AddCode "function getProperty(jsonObj, propertyName) { return jsonObj[propertyName]; } "
    ScriptEngine.AddCode "function getKeys(jsonObj) { var keys = new Array(); for (var i in jsonObj) { keys.push(i); } return keys; } "
#End If
End Sub

Public Function DecodeJsonString(ByVal JsonString As String)
#If Mac Then
    Set DecodeJsonString = JsonConverter.ParseJson(JsonString)
#Else
    Set DecodeJsonString = ScriptEngine.eval("(" + JsonString + ")")
#End If
End Function

Public Function GetProperty(ByVal jsonObject As Object, ByVal propertyName As String) As Variant
#If Mac Then
    Dim child As Object, value As Variant
    If smfJSONPortableProperty(jsonObject, propertyName, child, value) Then
        If child Is Nothing Then
            GetProperty = value
        Else
            Set GetProperty = child
        End If
    End If
#Else
    GetProperty = ScriptEngine.Run("getProperty", jsonObject, propertyName)
#End If
End Function

Public Function GetObjectProperty(ByVal jsonObject As Object, ByVal propertyName As String) As Object
#If Mac Then
    Dim child As Object, value As Variant
    If smfJSONPortableProperty(jsonObject, propertyName, child, value) Then
        If Not child Is Nothing Then
            Set GetObjectProperty = child
            Exit Function
        End If
    End If
    Err.Raise 13, "GetObjectProperty", "JSON property is not an object or array: " & propertyName
#Else
    Set GetObjectProperty = ScriptEngine.Run("getProperty", jsonObject, propertyName)
#End If
End Function

Public Function GetKeys(ByVal jsonObject As Object) As String()
    Dim Length As Long
    Dim KeysArray() As String
    Dim KeysObject As Object
    Dim Index As Long
    Dim Key As Variant

#If Mac Then
    If jsonObject Is Nothing Then Err.Raise 5, "GetKeys", "JSON node is missing"
    If TypeName(jsonObject) <> "Dictionary" And TypeName(jsonObject) <> "Collection" Then
        Err.Raise 5, "GetKeys", "JSON node is not an object or array"
    End If
    If jsonObject.Count = 0 Then
        KeysArray = Split(vbNullString)
    Else
        ReDim KeysArray(0 To jsonObject.Count - 1)
        If TypeName(jsonObject) = "Collection" Then
            For Index = 0 To jsonObject.Count - 1
                KeysArray(Index) = CStr(Index)
            Next Index
        Else
            Dim objectKeys As Variant
            objectKeys = jsonObject.Keys
            For Index = 0 To jsonObject.Count - 1
                KeysArray(Index) = CStr(objectKeys(Index))
            Next Index
        End If
    End If
#Else
    Set KeysObject = ScriptEngine.Run("getKeys", jsonObject)
    Length = GetProperty(KeysObject, "length")
    ReDim KeysArray(Length - 1)
    Index = 0
    For Each Key In KeysObject
        KeysArray(Index) = Key
        Index = Index + 1
    Next
#End If
    GetKeys = KeysArray
End Function

Function smfJSONExtractField(pJSONData As String, pJSONFields As String)
    '-----------------------------------------------------------------------------------
    ' 08/28/23 - Function runs in Excel x32 only.
    ' Native Mac also uses this entry point with JsonConverter.
    '
    '-----------------------------------------------------------------------------------
#If Mac Then
    Dim child As Object, value As Variant
    If smfJSONPortableNode(pJSONData, pJSONFields, child, value) Then
        If child Is Nothing Then
            smfJSONExtractField = value
        Else
            Set smfJSONExtractField = child
        End If
    Else
        smfJSONExtractField = "Not Found"
    End If
#Else
    Dim a1 As Variant, i1 As Integer, oJSON As Object

    InitScriptEngine
    
    a1 = Split(pJSONFields, ".")
    
    Set oJSON = DecodeJsonString(CStr(pJSONData))
    For i1 = 0 To UBound(a1) - 1
        Set oJSON = GetObjectProperty(oJSON, a1(i1))
        Next i1
        
    smfJSONExtractField = GetProperty(oJSON, a1(i1))
    If smfJSONExtractField = Empty Then smfJSONExtractField = "Not Found"
#End If

End Function
        
Function smfJSONExtractField_x64(pJSONData As String, pFieldNames As String)
    '-----------------------------------------------------------------------------------
    ' 08/28/23 - Created for Excel x64
    '            Uses "JSONConverter" to parse thru JSON string.
    '-----------------------------------------------------------------------------------

#If Mac Then
    Dim child As Object, value As Variant
    If smfJSONPortableNode(pJSONData, pFieldNames, child, value) Then
        If child Is Nothing Then
            smfJSONExtractField_x64 = value
        Else
            Set smfJSONExtractField_x64 = child
        End If
    Else
        smfJSONExtractField_x64 = "Not Found"
    End If
#Else
    Dim aFieldNames As Variant, i1 As Integer, oJSON As Object
    
    Dim vJSONKey0 As Variant, vJSONKey1 As Variant, vJSONKey2 As Variant, vJSONKey3 As Variant, vJSONKey4 As Variant
    Dim vJSONKey5 As Variant, vJSONKey6 As Variant, vJSONKey7 As Variant, vJSONKey8 As Variant, vJSONKey9 As Variant
    Dim vJSONFieldValue As Variant
    
    aFieldNames = Split(pFieldNames, ".")
    
    Set oJSON = JsonConverter.ParseJson(pJSONData)
       
    For i1 = 0 To UBound(aFieldNames)
        Select Case i1
            Case 0: vJSONKey0 = smfChkJSONFieldType(aFieldNames(i1))
            Case 1: vJSONKey1 = smfChkJSONFieldType(aFieldNames(i1))
            Case 2: vJSONKey2 = smfChkJSONFieldType(aFieldNames(i1))
            Case 3: vJSONKey3 = smfChkJSONFieldType(aFieldNames(i1))
            Case 4: vJSONKey4 = smfChkJSONFieldType(aFieldNames(i1))
            Case 5: vJSONKey5 = smfChkJSONFieldType(aFieldNames(i1))
            Case 6: vJSONKey6 = smfChkJSONFieldType(aFieldNames(i1))
            Case 7: vJSONKey7 = smfChkJSONFieldType(aFieldNames(i1))
            Case 8: vJSONKey8 = smfChkJSONFieldType(aFieldNames(i1))
            Case 9: vJSONKey9 = smfChkJSONFieldType(aFieldNames(i1))
        End Select
    Next
            
    Select Case i1
        Case 10: vJSONFieldValue = oJSON(vJSONKey0)(vJSONKey1)(vJSONKey2)(vJSONKey3)(vJSONKey4)(vJSONKey5)(vJSONKey6)(vJSONKey7)(vJSONKey8)(vJSONKey9)
        Case 9: vJSONFieldValue = oJSON(vJSONKey0)(vJSONKey1)(vJSONKey2)(vJSONKey3)(vJSONKey4)(vJSONKey5)(vJSONKey6)(vJSONKey7)(vJSONKey8)
        Case 8: vJSONFieldValue = oJSON(vJSONKey0)(vJSONKey1)(vJSONKey2)(vJSONKey3)(vJSONKey4)(vJSONKey5)(vJSONKey6)(vJSONKey7)
        Case 7: vJSONFieldValue = oJSON(vJSONKey0)(vJSONKey1)(vJSONKey2)(vJSONKey3)(vJSONKey4)(vJSONKey5)(vJSONKey6)
        Case 6: vJSONFieldValue = oJSON(vJSONKey0)(vJSONKey1)(vJSONKey2)(vJSONKey3)(vJSONKey4)(vJSONKey5)
        Case 5: vJSONFieldValue = oJSON(vJSONKey0)(vJSONKey1)(vJSONKey2)(vJSONKey3)(vJSONKey4)
        Case 4: vJSONFieldValue = oJSON(vJSONKey0)(vJSONKey1)(vJSONKey2)(vJSONKey3)
        Case Else
            vJSONFieldValue = "JSON key field (" & i1 + 1 & ") out of range"
    End Select
    
    If vJSONFieldValue = Empty Then vJSONFieldValue = "Not Found"
    
    smfJSONExtractField_x64 = vJSONFieldValue
#End If

End Function
    
Function smfChkJSONFieldType(pField)
    '-----------------------------------------------------------------------------------
    ' 08/28/23 - Created for Excel x64
    '            "JSONConverter" arrays start at 1 not 0.
    '            "Numeric" key fields also need to be numeric type, not 'variant'
    '-----------------------------------------------------------------------------------
    Dim vField As Variant
    
    If IsNumeric(pField) Then
        vField = pField + 1             ' Also converts to numeric type
    Else
        vField = pField
    End If
           
    smfChkJSONFieldType = vField
    
End Function

Function smfJSONExtractKeys(pJSONData As String, pFieldName As String)
    Dim a1 As Variant, i1 As Long, o1 As Object, oJSON As Object, v1 As Variant, vReturn As Variant

#If Mac Then
    Dim value As Variant
    If Not smfJSONPortableNode(pJSONData, pFieldName, oJSON, value) Then
        Err.Raise 5, "smfJSONExtractKeys", "JSON path was not found: " & pFieldName
    End If
    If oJSON Is Nothing Then
        Err.Raise 5, "smfJSONExtractKeys", "JSON path must identify an object or array"
    End If
#Else
    InitScriptEngine
    
    a1 = Split(pFieldName, ".")
    
    Set oJSON = DecodeJsonString(CStr(pJSONData))
    For i1 = 0 To UBound(a1)
        Set oJSON = GetObjectProperty(oJSON, a1(i1))
        Next i1
#End If
        
    v1 = GetKeys(oJSON)
    ReDim vReturn(1 To UBound(v1) + 2, 1 To 1)
    For i1 = 0 To UBound(v1)
        vReturn(i1 + 1, 1) = v1(i1)
        Next i1
    vReturn(i1 + 1, 1) = "#N/A"
    smfJSONExtractKeys = vReturn
    
End Function


#If Mac Then
' JSON path array indexes are zero-based. Numeric object keys remain strings.
' An empty path selects the root object or array. Dots in property names are
' not escaped by the existing SMF path format.
Private Function smfJSONPortableNode(ByVal pJSONData As String, ByVal pPath As String, _
                                    ByRef pNode As Object, ByRef pValue As Variant) As Boolean
    Dim currentNode As Object, parts As Variant, i As Long
    Set pNode = Nothing
    pValue = Empty
    Set currentNode = JsonConverter.ParseJson(pJSONData)
    If Len(pPath) = 0 Then
        Set pNode = currentNode
        smfJSONPortableNode = True
        Exit Function
    End If

    parts = Split(pPath, ".")
    For i = LBound(parts) To UBound(parts)
        If Not smfJSONPortableProperty(currentNode, CStr(parts(i)), pNode, pValue) Then Exit Function
        If i < UBound(parts) Then
            If pNode Is Nothing Then Exit Function
            Set currentNode = pNode
        End If
    Next i
    smfJSONPortableNode = True
End Function

Private Function smfJSONPortableProperty(ByVal pNode As Object, ByVal pName As String, _
                                        ByRef pChild As Object, ByRef pValue As Variant) As Boolean
    Dim key As Variant, i As Long, arrayIndex As Double
    Set pChild = Nothing
    pValue = Empty
    If pNode Is Nothing Then Exit Function

    Select Case TypeName(pNode)
        Case "Dictionary"
            key = pName
            If Not pNode.Exists(key) Then Exit Function
        Case "Collection"
            If Len(pName) = 0 Then Exit Function
            If Len(pName) > 1 And Left$(pName, 1) = "0" Then Exit Function
            For i = 1 To Len(pName)
                If Mid$(pName, i, 1) < "0" Or Mid$(pName, i, 1) > "9" Then Exit Function
            Next i
            ' Check the range before converting to Long; huge indexes are absent.
            On Error GoTo MissingIndex
            arrayIndex = CDbl(pName)
            If arrayIndex >= pNode.Count Then Exit Function
            key = CLng(arrayIndex) + 1
            On Error GoTo 0
        Case Else
            Exit Function
    End Select

    If IsObject(pNode.Item(key)) Then
        Set pChild = pNode.Item(key)
    Else
        pValue = pNode.Item(key)
    End If
    smfJSONPortableProperty = True
    Exit Function
MissingIndex:
    Err.Clear
    smfJSONPortableProperty = False
End Function
#End If

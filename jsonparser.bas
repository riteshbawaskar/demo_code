Option Explicit
'''
' usage
'Sub Demo()
'    Dim json As String
'    json = "{""trade"":{""price"":123.45,""qty"":10},""trades"":[{""price"":100},{""price"":200}],""weird key"":{""p-x"":42}}"
'    
'    Dim root As Variant
'    root = JsonParse(json)                       ' Parse once
'
'    Debug.Print JsonPathGet(root, "$.trade.price")          ' 123.45
'    Debug.Print JsonPathGet(root, "$.trades[0].price")      ' 100
'End Sub
'
'


' ===== Public API =====
Public Function JsonParse(ByVal s As String) As Variant
    Dim p As Long: p = 1
    SkipWS s, p
    JsonParse = ParseValue(s, p)
    SkipWS s, p
    If p <= Len(s) Then Err.Raise 5, , "Unexpected trailing characters at position " & p
End Function

' JSONPath: $.a.b, $.arr[0].name, $.['key with spaces']
Public Function JsonPathGet(ByVal root As Variant, ByVal path As String) As Variant
    Dim p As Long: p = 1
    SkipWS path, p
    If p > Len(path) Or Mid$(path, p, 1) <> "$" Then Err.Raise 5, , "Path must start with $"
    p = p + 1
    Dim cur As Variant: cur = root
    
    Do While p <= Len(path)
        SkipWS path, p
        If p > Len(path) Then Exit Do
        
        Dim ch As String: ch = Mid$(path, p, 1)
        Select Case ch
            Case "."
                p = p + 1
                cur = PathReadProperty(path, p, cur)
            Case "["
                p = p + 1
                cur = PathReadBracket(path, p, cur) ' advances past ]
            Case Else
                Exit Do
        End Select
    Loop
    
    JsonPathGet = cur
End Function

' ======= Internal: JSON parser =======

Private Function ParseValue(ByRef s As String, ByRef p As Long) As Variant
    SkipWS s, p
    If p > Len(s) Then Err.Raise 5, , "Unexpected end of input"
    
    Dim ch As String: ch = Mid$(s, p, 1)
    Select Case ch
        Case "{": ParseValue = ParseObject(s, p)
        Case "[": ParseValue = ParseArray(s, p)
        Case """": ParseValue = ParseString(s, p)
        Case "t", "f": ParseValue = ParseBoolean(s, p)
        Case "n": ParseValue = ParseNull(s, p)
        Case "-", "0" To "9": ParseValue = ParseNumber(s, p)
        Case Else
            Err.Raise 5, , "Unexpected character '" & ch & "' at position " & p
    End Select
End Function

Private Function ParseObject(ByRef s As String, ByRef p As Long) As Variant
    ' Assumes s[p] = "{"
    p = p + 1
    SkipWS s, p
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    
    If p <= Len(s) And Mid$(s, p, 1) = "}" Then
        p = p + 1
        Set ParseObject = dict
        Exit Function
    End If
    
    Do
        SkipWS s, p
        If p > Len(s) Or Mid$(s, p, 1) <> """" Then Err.Raise 5, , "Expected string key at position " & p
        Dim key As String: key = ParseString(s, p)
        
        SkipWS s, p
        If p > Len(s) Or Mid$(s, p, 1) <> ":" Then Err.Raise 5, , "Expected ':' after object key at position " & p
        p = p + 1
        
        Dim val As Variant
        SkipWS s, p
        val = ParseValue(s, p)
        dict(key) = val
        
        SkipWS s, p
        If p > Len(s) Then Err.Raise 5, , "Unterminated object"
        Dim ch As String: ch = Mid$(s, p, 1)
        If ch = "}" Then
            p = p + 1
            Exit Do
        ElseIf ch = "," Then
            p = p + 1
        Else
            Err.Raise 5, , "Expected ',' or '}' at position " & p
        End If
    Loop
    
    Set ParseObject = dict
End Function

Private Function ParseArray(ByRef s As String, ByRef p As Long) As Variant
    ' Assumes s[p] = "["
    p = p + 1
    SkipWS s, p
    Dim col As New Collection
    
    If p <= Len(s) And Mid$(s, p, 1) = "]" Then
        p = p + 1
        Set ParseArray = col
        Exit Function
    End If
    
    Do
        Dim v As Variant
        v = ParseValue(s, p)
        col.Add v
        
        SkipWS s, p
        If p > Len(s) Then Err.Raise 5, , "Unterminated array"
        Dim ch As String: ch = Mid$(s, p, 1)
        If ch = "]" Then
            p = p + 1
            Exit Do
        ElseIf ch = "," Then
            p = p + 1
            SkipWS s, p
        Else
            Err.Raise 5, , "Expected ',' or ']' at position " & p
        End If
    Loop
    
    Set ParseArray = col
End Function

Private Function ParseString(ByRef s As String, ByRef p As Long) As String
    ' Assumes s[p] = """
    p = p + 1
    Dim sb As String
    Do While p <= Len(s)
        Dim ch As String: ch = Mid$(s, p, 1)
        p = p + 1
        If ch = """" Then Exit Do
        If ch = "\" Then
            If p > Len(s) Then Err.Raise 5, , "Invalid escape at end of string"
            Dim esc As String: esc = Mid$(s, p, 1)
            p = p + 1
            Select Case esc
                Case """": sb = sb & """"
                Case "\"": sb = sb & "\"
                Case "/":  sb = sb & "/"
                Case "b":  sb = sb & vbBack
                Case "f":  sb = sb & vbFormFeed
                Case "n":  sb = sb & vbLf
                Case "r":  sb = sb & vbCr
                Case "t":  sb = sb & vbTab
                Case "u"
                    If p + 3 > Len(s) Then Err.Raise 5, , "Invalid \u escape"
                    Dim hex4 As String: hex4 = Mid$(s, p, 4)
                    If Not IsHex4(hex4) Then Err.Raise 5, , "Invalid \u" & hex4
                    sb = sb & ChrW$(CLng("&H" & hex4))
                    p = p + 4
                Case Else
                    Err.Raise 5, , "Invalid escape '\\" & esc & "'"
            End Select
        Else
            sb = sb & ch
        End If
    Loop
    If p > Len(s) And Right$(sb, 1) <> """" Then Err.Raise 5, , "Unterminated string"
    ParseString = sb
End Function

Private Function ParseNumber(ByRef s As String, ByRef p As Long) As Double
    Dim startP As Long: startP = p
    Dim ch As String
    
    ' sign
    If Mid$(s, p, 1) = "-" Then p = p + 1
    ' int
    If p > Len(s) Then Err.Raise 5, , "Invalid number"
    ch = Mid$(s, p, 1)
    If ch = "0" Then
        p = p + 1
    ElseIf ch >= "1" And ch <= "9" Then
        Do While p <= Len(s)
            ch = Mid$(s, p, 1)
            If ch < "0" Or ch > "9" Then Exit Do
            p = p + 1
        Loop
    Else
        Err.Raise 5, , "Invalid number at position " & startP
    End If
    ' frac
    If p <= Len(s) And Mid$(s, p, 1) = "." Then
        p = p + 1
        If p > Len(s) Or Mid$(s, p, 1) < "0" Or Mid$(s, p, 1) > "9" Then Err.Raise 5, , "Invalid fraction"
        Do While p <= Len(s)
            ch = Mid$(s, p, 1)
            If ch < "0" Or ch > "9" Then Exit Do
            p = p + 1
        Loop
    End If
    ' exp
    If p <= Len(s) Then
        ch = Mid$(s, p, 1)
        If ch = "e" Or ch = "E" Then
            p = p + 1
            If p <= Len(s) Then
                ch = Mid$(s, p, 1)
                If ch = "+" Or ch = "-" Then p = p + 1
            End If
            If p > Len(s) Or Mid$(s, p, 1) < "0" Or Mid$(s, p, 1) > "9" Then Err.Raise 5, , "Invalid exponent"
            Do While p <= Len(s)
                ch = Mid$(s, p, 1)
                If ch < "0" Or ch > "9" Then Exit Do
                p = p + 1
            Loop
        End If
    End If
    
    ParseNumber = CDbl(Mid$(s, startP, p - startP))
End Function

Private Function ParseBoolean(ByRef s As String, ByRef p As Long) As Boolean
    If Mid$(s, p, 4) = "true" Then
        p = p + 4
        ParseBoolean = True
    ElseIf Mid$(s, p, 5) = "false" Then
        p = p + 5
        ParseBoolean = False
    Else
        Err.Raise 5, , "Invalid boolean at position " & p
    End If
End Function

Private Function ParseNull(ByRef s As String, ByRef p As Long) As Variant
    If Mid$(s, p, 4) <> "null" Then Err.Raise 5, , "Invalid null at position " & p
    p = p + 4
    ParseNull = Null
End Function

Private Sub SkipWS(ByRef s As String, ByRef p As Long)
    Do While p <= Len(s)
        Select Case AscW(Mid$(s, p, 1))
            Case 9, 10, 13, 32 ' \t \n \r space
                p = p + 1
            Case Else
                Exit Sub
        End Select
    Loop
End Sub

Private Function IsHex4(ByVal s As String) As Boolean
    Dim i As Long, ch As Integer
    If Len(s) <> 4 Then Exit Function
    For i = 1 To 4
        ch = AscW(Mid$(s, i, 1))
        If Not ((ch >= 48 And ch <= 57) Or (ch >= 65 And ch <= 70) Or (ch >= 97 And ch <= 102)) Then Exit Function
    Next
    IsHex4 = True
End Function

' ======= Internal: JSONPath pieces =======

Private Function PathReadProperty(ByRef path As String, ByRef p As Long, ByVal cur As Variant) As Variant
    SkipWS path, p
    If p > Len(path) Then Err.Raise 5, , "Expected property after '.'"
    
    Dim ch As String: ch = Mid$(path, p, 1)
    Dim name As String
    
    If ch = "[" Then
        Err.Raise 5, , "Unexpected '[' after '.'; use $.prop[index]"
    ElseIf ch = "'" Or ch = """" Then
        name = PathQuotedIdentifier(path, p)
    Else
        name = PathBareIdentifier(path, p)
    End If
    
    PathReadProperty = GetObjectProperty(cur, name)
End Function

Private Function PathReadBracket(ByRef path As String, ByRef p As Long, ByVal cur As Variant) As Variant
    SkipWS path, p
    If p > Len(path) Then Err.Raise 5, , "Unterminated '['"
    
    Dim ch As String: ch = Mid$(path, p, 1)
    Dim result As Variant
    
    If ch = """" Or ch = "'" Then
        Dim name As String: name = PathQuotedIdentifier(path, p)
        SkipWS path, p
        If p > Len(path) Or Mid$(path, p, 1) <> "]" Then Err.Raise 5, , "Expected ']' after string key"
        p = p + 1
        result = GetObjectProperty(cur, name)
    Else
        ' index
        Dim sign As Long: sign = 1
        If ch = "-" Then sign = -1: p = p + 1
        Dim numStart As Long: numStart = p
        Do While p <= Len(path) And Mid$(path, p, 1) Like "[0-9]"
            p = p + 1
        Loop
        If numStart = p Then Err.Raise 5, , "Expected array index"
        Dim idx As Long: idx = CLng(Mid$(path, numStart, p - numStart)) * sign
        SkipWS path, p
        If p > Len(path) Or Mid$(path, p, 1) <> "]" Then Err.Raise 5, , "Expected ']' after index"
        p = p + 1
        result = GetArrayIndex(cur, idx)
    End If
    
    PathReadBracket = result
End Function

Private Function PathBareIdentifier(ByRef path As String, ByRef p As Long) As String
    Dim startP As Long: startP = p
    Do While p <= Len(path)
        Dim ch As String: ch = Mid$(path, p, 1)
        If (ch Like "[A-Za-z0-9_-$]") Then
            p = p + 1
        Else
            Exit Do
        End If
    Loop
    If startP = p Then Err.Raise 5, , "Expected identifier at position " & p
    PathBareIdentifier = Mid$(path, startP, p - startP)
End Function

Private Function PathQuotedIdentifier(ByRef path As String, ByRef p As Long) As String
    Dim quote As String: quote = Mid$(path, p, 1)
    If quote <> """" And quote <> "'" Then Err.Raise 5, , "Expected quote"
    p = p + 1
    Dim sb As String
    Do While p <= Len(path)
        Dim ch As String: ch = Mid$(path, p, 1): p = p + 1
        If ch = quote Then Exit Do
        If ch = "\" Then
            If p > Len(path) Then Err.Raise 5, , "Invalid escape in identifier"
            Dim esc As String: esc = Mid$(path, p, 1): p = p + 1
            If esc = quote Then
                sb = sb & quote
            ElseIf esc = "\" Then
                sb = sb & "\"
            Else
                sb = sb & esc ' keep other escapes literal for identifiers
            End If
        Else
            sb = sb & ch
        End If
    Loop
    PathQuotedIdentifier = sb
End Function

Private Function GetObjectProperty(ByVal obj As Variant, ByVal name As String) As Variant
    If IsObject(obj) Then
        If TypeName(obj) = "Dictionary" Or TypeName(obj) = "Scripting.Dictionary" Then
            Dim d As Object: Set d = obj
            If d.Exists(name) Then
                GetObjectProperty = d(name)
            Else
                Err.Raise 5, , "Property '" & name & "' not found"
            End If
        Else
            Err.Raise 5, , "Not an object for property access"
        End If
    Else
        Err.Raise 5, , "Primitive value has no properties"
    End If
End Function

' JSON arrays are 0-based in path; VBA Collections are 1-based
Private Function GetArrayIndex(ByVal arr As Variant, ByVal idx As Long) As Variant
    If IsObject(arr) And TypeName(arr) = "Collection" Then
        Dim c As Collection: Set c = arr
        Dim vIdx As Long: vIdx = idx + 1
        If vIdx < 1 Or vIdx > c.Count Then Err.Raise 5, , "Array index out of bounds"
        GetArrayIndex = c.Item(vIdx)
    Else
        Err.Raise 5, , "Not an array for index access"
    End If
End Function

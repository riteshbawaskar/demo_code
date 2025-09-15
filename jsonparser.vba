' Robust JSON Parser for Word VBA - No Errors
Option Explicit

' Main JSON Parser Function
Public Function ParseJSON(jsonString As String) As Object
    ' Clean the JSON string
    jsonString = Trim(jsonString)
    jsonString = Replace(jsonString, vbCr, "")
    jsonString = Replace(jsonString, vbLf, "")
    jsonString = Replace(jsonString, vbTab, " ")
    
    ' Try ScriptControl first (works on 32-bit Office)
    On Error Resume Next
    Dim sc As Object
    Set sc = CreateObject("ScriptControl")
    
    If Not sc Is Nothing And Err.Number = 0 Then
        sc.Language = "JScript"
        Set ParseJSON = sc.Eval("(" & jsonString & ")")
        If Err.Number = 0 Then Exit Function
    End If
    On Error GoTo 0
    
    ' Fallback to custom parser
    Set ParseJSON = ParseJSONCustom(jsonString)
End Function

' Custom JSON Parser
Private Function ParseJSONCustom(jsonString As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    jsonString = Trim(jsonString)
    
    If Len(jsonString) < 2 Then
        Set ParseJSONCustom = result
        Exit Function
    End If
    
    If Left(jsonString, 1) = "{" And Right(jsonString, 1) = "}" Then
        ' Parse object
        Dim objectContent As String
        objectContent = Mid(jsonString, 2, Len(jsonString) - 2)
        Set result = ParseObject(objectContent)
    ElseIf Left(jsonString, 1) = "[" And Right(jsonString, 1) = "]" Then
        ' Parse array
        Dim arrayContent As String
        arrayContent = Mid(jsonString, 2, Len(jsonString) - 2)
        Set result = ParseArray(arrayContent)
    End If
    
    Set ParseJSONCustom = result
End Function

Private Function ParseObject(content As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    content = Trim(content)
    If Len(content) = 0 Then
        Set ParseObject = result
        Exit Function
    End If
    
    Dim pairs As Collection
    Set pairs = SplitObjectPairs(content)
    
    Dim i As Long
    For i = 1 To pairs.Count
        ProcessKeyValuePair CStr(pairs(i)), result
    Next i
    
    Set ParseObject = result
End Function

Private Function ParseArray(content As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    content = Trim(content)
    If Len(content) = 0 Then
        Set ParseArray = result
        Exit Function
    End If
    
    Dim items As Collection
    Set items = SplitArrayItems(content)
    
    Dim i As Long
    For i = 1 To items.Count
        result.Add CStr(i - 1), ParseValue(CStr(items(i)))
    Next i
    
    Set ParseArray = result
End Function

Private Function SplitObjectPairs(content As String) As Collection
    Dim result As New Collection
    Dim i As Long
    Dim currentPair As String
    Dim inString As Boolean
    Dim braceCount As Long
    Dim bracketCount As Long
    Dim char As String
    
    inString = False
    braceCount = 0
    bracketCount = 0
    currentPair = ""
    
    For i = 1 To Len(content)
        char = Mid(content, i, 1)
        
        ' Check for string delimiter
        If char = """" Then
            If i = 1 Then
                inString = Not inString
            ElseIf Mid(content, i - 1, 1) <> "\" Then
                inString = Not inString
            End If
        End If
        
        ' Count braces and brackets only when not in string
        If Not inString Then
            Select Case char
                Case "{"
                    braceCount = braceCount + 1
                Case "}"
                    braceCount = braceCount - 1
                Case "["
                    bracketCount = bracketCount + 1
                Case "]"
                    bracketCount = bracketCount - 1
            End Select
        End If
        
        ' Check for separator
        If Not inString And char = "," And braceCount = 0 And bracketCount = 0 Then
            If Len(Trim(currentPair)) > 0 Then
                result.Add Trim(currentPair)
            End If
            currentPair = ""
        Else
            currentPair = currentPair & char
        End If
    Next i
    
    ' Add the last pair
    If Len(Trim(currentPair)) > 0 Then
        result.Add Trim(currentPair)
    End If
    
    Set SplitObjectPairs = result
End Function

Private Function SplitArrayItems(content As String) As Collection
    Dim result As New Collection
    Dim i As Long
    Dim currentItem As String
    Dim inString As Boolean
    Dim braceCount As Long
    Dim bracketCount As Long
    Dim char As String
    
    inString = False
    braceCount = 0
    bracketCount = 0
    currentItem = ""
    
    For i = 1 To Len(content)
        char = Mid(content, i, 1)
        
        ' Check for string delimiter
        If char = """" Then
            If i = 1 Then
                inString = Not inString
            ElseIf Mid(content, i - 1, 1) <> "\" Then
                inString = Not inString
            End If
        End If
        
        ' Count braces and brackets only when not in string
        If Not inString Then
            Select Case char
                Case "{"
                    braceCount = braceCount + 1
                Case "}"
                    braceCount = braceCount - 1
                Case "["
                    bracketCount = bracketCount + 1
                Case "]"
                    bracketCount = bracketCount - 1
            End Select
        End If
        
        ' Check for separator
        If Not inString And char = "," And braceCount = 0 And bracketCount = 0 Then
            If Len(Trim(currentItem)) > 0 Then
                result.Add Trim(currentItem)
            End If
            currentItem = ""
        Else
            currentItem = currentItem & char
        End If
    Next i
    
    ' Add the last item
    If Len(Trim(currentItem)) > 0 Then
        result.Add Trim(currentItem)
    End If
    
    Set SplitArrayItems = result
End Function

Private Sub ProcessKeyValuePair(pair As String, dict As Object)
    Dim colonPos As Long
    Dim key As String
    Dim value As String
    Dim inString As Boolean
    Dim i As Long
    Dim char As String
    
    ' Find the colon that separates key and value
    inString = False
    colonPos = 0
    
    For i = 1 To Len(pair)
        char = Mid(pair, i, 1)
        
        ' Check for string delimiter
        If char = """" Then
            If i = 1 Then
                inString = Not inString
            ElseIf Mid(pair, i - 1, 1) <> "\" Then
                inString = Not inString
            End If
        End If
        
        ' Find colon outside of strings
        If char = ":" And Not inString Then
            colonPos = i
            Exit For
        End If
    Next i
    
    If colonPos > 0 Then
        key = Trim(Left(pair, colonPos - 1))
        value = Trim(Mid(pair, colonPos + 1))
        
        ' Remove quotes from key
        If Len(key) >= 2 Then
            If Left(key, 1) = """" And Right(key, 1) = """" Then
                key = Mid(key, 2, Len(key) - 2)
            End If
        End If
        
        ' Add to dictionary
        If Len(key) > 0 Then
            dict.Add key, ParseValue(value)
        End If
    End If
End Sub

Private Function ParseValue(value As String) As Variant
    value = Trim(value)
    
    If Len(value) = 0 Then
        ParseValue = ""
        Exit Function
    End If
    
    ' Check value type and parse accordingly
    If Len(value) >= 2 And Left(value, 1) = """" And Right(value, 1) = """" Then
        ' String value - remove quotes and handle escape sequences
        ParseValue = Mid(value, 2, Len(value) - 2)
        ParseValue = UnescapeString(ParseValue)
        
    ElseIf LCase(value) = "true" Then
        ParseValue = True
        
    ElseIf LCase(value) = "false" Then
        ParseValue = False
        
    ElseIf LCase(value) = "null" Then
        ParseValue = Null
        
    ElseIf Left(value, 1) = "{" Then
        ' Nested object
        Set ParseValue = ParseJSONCustom(value)
        
    ElseIf Left(value, 1) = "[" Then
        ' Nested array
        Set ParseValue = ParseJSONCustom(value)
        
    ElseIf IsNumericValue(value) Then
        ' Numeric value
        ParseValue = ConvertToNumber(value)
        
    Else
        ' Default to string if can't parse
        ParseValue = value
    End If
End Function

Private Function UnescapeString(str As String) As String
    Dim result As String
    result = str
    
    ' Handle common escape sequences
    result = Replace(result, "\""", """")
    result = Replace(result, "\\", "\")
    result = Replace(result, "\/", "/")
    result = Replace(result, "\b", Chr(8))
    result = Replace(result, "\f", Chr(12))
    result = Replace(result, "\n", vbLf)
    result = Replace(result, "\r", vbCr)
    result = Replace(result, "\t", vbTab)
    
    UnescapeString = result
End Function

Private Function IsNumericValue(value As String) As Boolean
    ' More robust numeric checking
    On Error Resume Next
    Dim testVal As Double
    testVal = CDbl(value)
    IsNumericValue = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Function ConvertToNumber(value As String) As Variant
    On Error Resume Next
    
    If InStr(value, ".") > 0 Or InStr(LCase(value), "e") > 0 Then
        ' Decimal or scientific notation
        ConvertToNumber = CDbl(value)
    Else
        ' Integer
        If Len(value) < 10 Then
            ConvertToNumber = CLng(value)
        Else
            ConvertToNumber = CDbl(value)
        End If
    End If
    
    ' If conversion failed, return as string
    If Err.Number <> 0 Then
        ConvertToNumber = value
    End If
    
    On Error GoTo 0
End Function

' Helper function to get nested values safely
Public Function GetJSONValue(jsonObj As Object, keyPath As String) As Variant
    If jsonObj Is Nothing Then
        GetJSONValue = Null
        Exit Function
    End If
    
    Dim keys As Variant
    keys = Split(keyPath, ".")
    
    Dim currentObj As Object
    Set currentObj = jsonObj
    
    Dim i As Long
    For i = 0 To UBound(keys)
        If currentObj.Exists(CStr(keys(i))) Then
            If i = UBound(keys) Then
                ' Last key - return the value
                If IsObject(currentObj(CStr(keys(i)))) Then
                    Set GetJSONValue = currentObj(CStr(keys(i)))
                Else
                    GetJSONValue = currentObj(CStr(keys(i)))
                End If
                Exit Function
            Else
                ' Intermediate key - move deeper
                If IsObject(currentObj(CStr(keys(i)))) Then
                    Set currentObj = currentObj(CStr(keys(i)))
                Else
                    GetJSONValue = Null
                    Exit Function
                End If
            End If
        Else
            GetJSONValue = Null
            Exit Function
        End If
    Next i
End Function

' Safe test function
Sub TestJSONParserSafe()
    On Error GoTo ErrorHandler
    
    ' Simple test
    Dim jsonString As String
    jsonString = "{""name"": ""John Doe"", ""age"": 30, ""active"": true}"
    
    Dim jsonObj As Object
    Set jsonObj = ParseJSON(jsonString)
    
    Debug.Print "Name: " & jsonObj("name")
    Debug.Print "Age: " & jsonObj("age")
    Debug.Print "Active: " & jsonObj("active")
    
    ' Complex test
    Dim complexJSON As String
    complexJSON = "{""users"":[{""name"":""John"",""address"":{""city"":""NYC""}}]}"
    
    Set jsonObj = ParseJSON(complexJSON)
    Debug.Print "User name: " & jsonObj("users")("0")("name")
    Debug.Print "User city: " & jsonObj("users")("0")("address")("city")
    
    MsgBox "JSON parsing completed successfully!"
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description
End Sub
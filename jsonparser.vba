' Simple JSON Parser for Word VBA - Fixed VBA Syntax
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
    
    If Left(jsonString, 1) = "{" And Right(jsonString, 1) = "}" Then
        ' Parse object
        Set result = ParseObject(Mid(jsonString, 2, Len(jsonString) - 2))
    ElseIf Left(jsonString, 1) = "[" And Right(jsonString, 1) = "]" Then
        ' Parse array
        Set result = ParseArray(Mid(jsonString, 2, Len(jsonString) - 2))
    End If
    
    Set ParseJSONCustom = result
End Function

Private Function ParseObject(content As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    If Len(Trim(content)) = 0 Then
        Set ParseObject = result
        Exit Function
    End If
    
    Dim pairs As Collection
    Set pairs = SplitObjectPairs(content)
    
    Dim i As Long
    For i = 1 To pairs.Count
        ProcessKeyValuePair pairs(i), result
    Next i
    
    Set ParseObject = result
End Function

Private Function ParseArray(content As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    If Len(Trim(content)) = 0 Then
        Set ParseArray = result
        Exit Function
    End If
    
    Dim items As Collection
    Set items = SplitArrayItems(content)
    
    Dim i As Long
    For i = 1 To items.Count
        result.Add CStr(i - 1), ParseValue(items(i))
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
        
        If char = """" And (i = 1 Or Mid(content, i - 1, 1) <> "\") Then
            inString = Not inString
        End If
        
        If Not inString Then
            If char = "{" Then braceCount = braceCount + 1
            If char = "}" Then braceCount = braceCount - 1
            If char = "[" Then bracketCount = bracketCount + 1
            If char = "]" Then bracketCount = bracketCount - 1
        End If
        
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
        
        If char = """" And (i = 1 Or Mid(content, i - 1, 1) <> "\") Then
            inString = Not inString
        End If
        
        If Not inString Then
            If char = "{" Then braceCount = braceCount + 1
            If char = "}" Then braceCount = braceCount - 1
            If char = "[" Then bracketCount = bracketCount + 1
            If char = "]" Then bracketCount = bracketCount - 1
        End If
        
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
    
    ' Find the colon that separates key and value
    inString = False
    colonPos = 0
    
    For i = 1 To Len(pair)
        If Mid(pair, i, 1) = """" And (i = 1 Or Mid(pair, i - 1, 1) <> "\") Then
            inString = Not inString
        End If
        If Mid(pair, i, 1) = ":" And Not inString Then
            colonPos = i
            Exit For
        End If
    Next i
    
    If colonPos > 0 Then
        key = Trim(Left(pair, colonPos - 1))
        value = Trim(Mid(pair, colonPos + 1))
        
        ' Remove quotes from key
        If Left(key, 1) = """" And Right(key, 1) = """" Then
            key = Mid(key, 2, Len(key) - 2)
        End If
        
        dict.Add key, ParseValue(value)
    End If
End Sub

Private Function ParseValue(value As String) As Variant
    value = Trim(value)
    
    If Len(value) = 0 Then
        ParseValue = Null
        Exit Function
    End If
    
    If Left(value, 1) = """" And Right(value, 1) = """" Then
        ' String value
        ParseValue = Mid(value, 2, Len(value) - 2)
        ParseValue = Replace(ParseValue, "\""", """")
        ParseValue = Replace(ParseValue, "\\", "\")
        ParseValue = Replace(ParseValue, "\/", "/")
        ParseValue = Replace(ParseValue, "\b", Chr(8))
        ParseValue = Replace(ParseValue, "\f", Chr(12))
        ParseValue = Replace(ParseValue, "\n", vbLf)
        ParseValue = Replace(ParseValue, "\r", vbCr)
        ParseValue = Replace(ParseValue, "\t", vbTab)
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
    ElseIf IsNumeric(value) Or (Left(value, 1) = "-" And IsNumeric(Mid(value, 2))) Then
        ' Numeric value
        If InStr(value, ".") > 0 Or InStr(LCase(value), "e") > 0 Then
            ParseValue = CDbl(value)
        Else
            If Len(value) < 10 Then
                ParseValue = CLng(value)
            Else
                ParseValue = CDbl(value)
            End If
        End If
    Else
        ' Default to string if can't parse
        ParseValue = value
    End If
End Function

' Helper function to get nested values
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
        If currentObj.Exists(keys(i)) Then
            If i = UBound(keys) Then
                ' Last key - return the value
                If IsObject(currentObj(keys(i))) Then
                    Set GetJSONValue = currentObj(keys(i))
                Else
                    GetJSONValue = currentObj(keys(i))
                End If
                Exit Function
            Else
                ' Intermediate key - move deeper
                If IsObject(currentObj(keys(i))) Then
                    Set currentObj = currentObj(keys(i))
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

' Example usage
Sub TestJSONParser()
    Dim jsonString As String
    jsonString = "{""name"": ""John Doe"", ""age"": 30, ""city"": ""New York"", " & _
                 """hobbies"": [""reading"", ""swimming""], " & _
                 """address"": {""street"": ""123 Main St"", ""zip"": ""10001""}}"
    
    Dim jsonObj As Object
    Set jsonObj = ParseJSON(jsonString)
    
    ' Test basic access
    Debug.Print "Name: " & jsonObj("name")
    Debug.Print "Age: " & jsonObj("age")
    Debug.Print "City: " & jsonObj("city")
    
    ' Test array access
    Debug.Print "First hobby: " & jsonObj("hobbies")("0")
    Debug.Print "Second hobby: " & jsonObj("hobbies")("1")
    
    ' Test nested object access
    Debug.Print "Street: " & jsonObj("address")("street")
    Debug.Print "Zip: " & jsonObj("address")("zip")
    
    ' Test helper function
    Debug.Print "Zip using helper: " & GetJSONValue(jsonObj, "address.zip")
    
    MsgBox "JSON parsing completed! Check Debug window for results."
End Sub

' Test with complex nested JSON
Sub TestComplexJSON()
    Dim complexJSON As String
    complexJSON = "{""users"":[{""id"":1,""name"":""John"",""address"":{""city"":""NYC"",""coords"":{""lat"":40.7,""lng"":-74.0}}}]}"
    
    Dim jsonObj As Object
    Set jsonObj = ParseJSON(complexJSON)
    
    ' Access deeply nested values
    Debug.Print "User name: " & jsonObj("users")("0")("name")
    Debug.Print "User city: " & jsonObj("users")("0")("address")("city")
    Debug.Print "User latitude: " & jsonObj("users")("0")("address")("coords")("lat")
    
    MsgBox "Complex JSON parsing completed!"
End Sub
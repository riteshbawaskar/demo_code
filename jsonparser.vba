' JSON Parser for Word VBA
' Add this to a module in Word VBA

Option Explicit

' Main JSON Parser Class
Public Function ParseJSON(jsonString As String) As Object
    ' Remove extra whitespace and line breaks
    jsonString = Trim(jsonString)
    jsonString = Replace(jsonString, vbCr, "")
    jsonString = Replace(jsonString, vbLf, "")
    
    ' Try using ScriptControl first (works on 32-bit Office)
    On Error Resume Next
    Dim sc As Object
    Set sc = CreateObject("ScriptControl")
    
    If Not sc Is Nothing Then
        sc.Language = "JScript"
        Set ParseJSON = sc.Eval("(" & jsonString & ")")
        Exit Function
    End If
    On Error GoTo 0
    
    ' Fallback to custom parser for 64-bit or if ScriptControl fails
    Set ParseJSON = ParseJSONCustom(jsonString)
End Function

' Custom JSON Parser (works on both 32-bit and 64-bit)
Private Function ParseJSONCustom(jsonString As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    jsonString = Trim(jsonString)
    
    If Left(jsonString, 1) = "{" And Right(jsonString, 1) = "}" Then
        ' Parse object
        Set result = ParseObject(Mid(jsonString, 2, Len(jsonString) - 2))
    ElseIf Left(jsonString, 1) = "[" And Right(jsonString, 1) = "]" Then
        ' Parse array - convert to dictionary with numeric keys
        Set result = ParseArray(Mid(jsonString, 2, Len(jsonString) - 2))
    End If
    
    Set ParseJSONCustom = result
End Function

Private Function ParseObject(content As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    Dim key As String
    Dim value As String
    Dim inString As Boolean
    Dim braceCount As Long
    Dim bracketCount As Long
    Dim currentPair As String
    Dim char As String
    
    i = 1
    inString = False
    braceCount = 0
    bracketCount = 0
    currentPair = ""
    
    Do While i <= Len(content)
        char = Mid(content, i, 1)
        
        If char = """" And (i = 1 Or Mid(content, i - 1, 1) <> "\") Then
            inString = Not inString
        End If
        
        If Not inString Then
            If char = "{" Then braceCount = braceCount + 1
            If char = "}" Then braceCount = braceCount - 1
            If char = "[" Then bracketCount = bracketCount + 1
            If char = "]" Then bracketCount = bracketCount - 1
            
            If char = "," And braceCount = 0 And bracketCount = 0 Then
                ' Process current pair
                ProcessKeyValuePair currentPair, result
                currentPair = ""
                i = i + 1
                Continue Do
            End If
        End If
        
        currentPair = currentPair & char
        i = i + 1
    Loop
    
    ' Process last pair
    If Len(currentPair) > 0 Then
        ProcessKeyValuePair currentPair, result
    End If
    
    Set ParseObject = result
End Function

Private Function ParseArray(content As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    Dim items As Collection
    Set items = SplitJSONArray(content)
    
    Dim i As Long
    For i = 1 To items.Count
        result.Add CStr(i - 1), ParseValue(items(i))
    Next i
    
    Set ParseArray = result
End Function

Private Sub ProcessKeyValuePair(pair As String, dict As Object)
    Dim colonPos As Long
    Dim key As String
    Dim value As String
    Dim inString As Boolean
    Dim i As Long
    
    ' Find the colon that separates key and value
    inString = False
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
    
    If Left(value, 1) = """" And Right(value, 1) = """" Then
        ' String value
        ParseValue = Mid(value, 2, Len(value) - 2)
        ParseValue = Replace(ParseValue, "\""", """")
        ParseValue = Replace(ParseValue, "\\", "\")
    ElseIf value = "true" Then
        ParseValue = True
    ElseIf value = "false" Then
        ParseValue = False
    ElseIf value = "null" Then
        ParseValue = Null
    ElseIf Left(value, 1) = "{" Then
        ' Nested object
        Set ParseValue = ParseJSONCustom(value)
    ElseIf Left(value, 1) = "[" Then
        ' Nested array
        Set ParseValue = ParseJSONCustom(value)
    ElseIf IsNumeric(value) Then
        ' Numeric value
        If InStr(value, ".") > 0 Then
            ParseValue = CDbl(value)
        Else
            ParseValue = CLng(value)
        End If
    Else
        ' Default to string
        ParseValue = value
    End If
End Function

Private Function SplitJSONArray(content As String) As Collection
    Dim result As New Collection
    Dim i As Long
    Dim currentItem As String
    Dim inString As Boolean
    Dim braceCount As Long
    Dim bracketCount As Long
    Dim char As String
    
    i = 1
    inString = False
    braceCount = 0
    bracketCount = 0
    currentItem = ""
    
    Do While i <= Len(content)
        char = Mid(content, i, 1)
        
        If char = """" And (i = 1 Or Mid(content, i - 1, 1) <> "\") Then
            inString = Not inString
        End If
        
        If Not inString Then
            If char = "{" Then braceCount = braceCount + 1
            If char = "}" Then braceCount = braceCount - 1
            If char = "[" Then bracketCount = bracketCount + 1
            If char = "]" Then bracketCount = bracketCount - 1
            
            If char = "," And braceCount = 0 And bracketCount = 0 Then
                If Len(Trim(currentItem)) > 0 Then
                    result.Add Trim(currentItem)
                End If
                currentItem = ""
                i = i + 1
                Continue Do
            End If
        End If
        
        currentItem = currentItem & char
        i = i + 1
    Loop
    
    ' Add last item
    If Len(Trim(currentItem)) > 0 Then
        result.Add Trim(currentItem)
    End If
    
    Set SplitJSONArray = result
End Function

' Helper function to get value from parsed JSON
Public Function GetJSONValue(jsonObj As Object, keyPath As String) As Variant
    Dim keys As Variant
    Dim currentObj As Object
    Dim i As Long
    
    keys = Split(keyPath, ".")
    Set currentObj = jsonObj
    
    For i = 0 To UBound(keys)
        If currentObj.Exists(keys(i)) Then
            If IsObject(currentObj(keys(i))) Then
                Set currentObj = currentObj(keys(i))
            Else
                If i = UBound(keys) Then
                    GetJSONValue = currentObj(keys(i))
                    Exit Function
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
    
    Set GetJSONValue = currentObj
End Function

' Example usage macro
Sub ExampleJSONUsage()
    Dim jsonString As String
    Dim jsonObj As Object
    
    ' Sample JSON string
    jsonString = "{""name"": ""John Doe"", ""age"": 30, ""city"": ""New York"", ""hobbies"": [""reading"", ""swimming""], ""address"": {""street"": ""123 Main St"", ""zip"": ""10001""}}"
    
    ' Parse JSON
    Set jsonObj = ParseJSON(jsonString)
    
    ' Access values
    Debug.Print "Name: " & jsonObj("name")
    Debug.Print "Age: " & jsonObj("age")
    Debug.Print "City: " & jsonObj("city")
    Debug.Print "First hobby: " & jsonObj("hobbies")("0")
    Debug.Print "Street: " & jsonObj("address")("street")
    
    ' Using helper function
    Debug.Print "Zip using helper: " & GetJSONValue(jsonObj, "address.zip")
    
    ' Insert into document
    Dim doc As Document
    Set doc = ActiveDocument
    
    doc.Content.InsertAfter "Name: " & jsonObj("name") & vbCrLf
    doc.Content.InsertAfter "Age: " & jsonObj("age") & vbCrLf
    doc.Content.InsertAfter "City: " & jsonObj("city") & vbCrLf
End Sub

' Function to read JSON from file
Sub ReadJSONFromFile()
    Dim fileName As String
    Dim jsonString As String
    Dim jsonObj As Object
    Dim fileNum As Integer
    
    ' Open file dialog
    fileName = Application.FileDialog(msoFileDialogFilePicker).Show
    If fileName <> 0 Then
        fileName = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(1)
        
        ' Read file content
        fileNum = FreeFile
        Open fileName For Input As fileNum
        jsonString = Input$(LOF(fileNum), fileNum)
        Close fileNum
        
        ' Parse and use JSON
        Set jsonObj = ParseJSON(jsonString)
        
        ' Process the JSON object as needed
        MsgBox "JSON file loaded successfully!"
    End If
End Sub

' Function to create JSON from Word table
Sub CreateJSONFromTable()
    Dim tbl As Table
    Dim jsonString As String
    Dim i As Long, j As Long
    Dim headers() As String
    Dim rowData As String
    
    If ActiveDocument.Tables.Count = 0 Then
        MsgBox "No tables found in document"
        Exit Sub
    End If
    
    Set tbl = ActiveDocument.Tables(1)
    
    ' Get headers from first row
    ReDim headers(1 To tbl.Columns.Count)
    For j = 1 To tbl.Columns.Count
        headers(j) = Trim(tbl.Cell(1, j).Range.Text)
        headers(j) = Replace(headers(j), Chr(13) & Chr(7), "") ' Remove table cell markers
    Next j
    
    ' Build JSON array
    jsonString = "["
    
    For i = 2 To tbl.Rows.Count ' Start from row 2 (skip headers)
        If i > 2 Then jsonString = jsonString & ","
        jsonString = jsonString & "{"
        
        For j = 1 To tbl.Columns.Count
            If j > 1 Then jsonString = jsonString & ","
            Dim cellValue As String
            cellValue = Trim(tbl.Cell(i, j).Range.Text)
            cellValue = Replace(cellValue, Chr(13) & Chr(7), "") ' Remove table cell markers
            cellValue = Replace(cellValue, """", "\""") ' Escape quotes
            
            jsonString = jsonString & """" & headers(j) & """:""" & cellValue & """"
        Next j
        
        jsonString = jsonString & "}"
    Next i
    
    jsonString = jsonString & "]"
    
    ' Insert JSON at end of document
    ActiveDocument.Content.InsertAfter vbCrLf & vbCrLf & "Generated JSON:" & vbCrLf & jsonString
End Sub

'Dim jsonObj As Object
'Set jsonObj = ParseJSON("{""name"": ""John"", ""age"": 30}")

' Access values
'Debug.Print jsonObj("name")  ' Outputs: John
'Debug.Print jsonObj("age")   ' Outputs: 30
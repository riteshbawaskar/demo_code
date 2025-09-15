' GitLab Test Case Generator - Clean, Fresh Build
' Compact ChatGPT-style design with consistent formatting throughout

Option Explicit

' Global variables
Dim ConfigTable As Table
Dim StatusCell As Range

' === MAIN ENTRY POINT ===
Sub StartGitLabGenerator()
    Call ClearDocument
    Call CreateInterface
    Call SetupShortcuts
    
    MsgBox ChrW(9733) & " GitLab Test Case Generator Ready!" & vbCrLf & vbCrLf & _
           "Use Alt+F8 to run macros or keyboard shortcuts!", vbInformation
End Sub

' === INTERFACE CREATION ===
Sub CreateInterface()
    Dim doc As Document
    Dim rng As Range
    
    Set doc = ActiveDocument
    Set rng = doc.Range
    
    Call SetDocumentStyle
    Call AddHeader(rng)
    Call AddConfigForm(rng)
    Call AddActionArea(rng)
    Call AddInstructions(rng)
End Sub

' Set document-wide styling
Sub SetDocumentStyle()
    With ActiveDocument
        .Range.Font.Name = "Segoe UI"
        .Range.Font.Size = 11
        .Range.Font.Color = RGB(52, 53, 65)
        
        With .PageSetup
            .TopMargin = InchesToPoints(0.5)
            .BottomMargin = InchesToPoints(0.5)
            .LeftMargin = InchesToPoints(0.6)
            .RightMargin = InchesToPoints(0.6)
        End With
    End With
End Sub

' Add header section
Sub AddHeader(rng As Range)
    ' Main title
    rng.text = ChrW(9658) & " GitLab Test Case Generator " & ChrW(9668) & vbCrLf
    With rng
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(52, 53, 65)
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .ParagraphFormat.SpaceAfter = 8
        .Collapse wdCollapseEnd
    End With
    
    ' Subtitle
    rng.text = "AI-powered test case generation from GitLab epics and issues" & vbCrLf & vbCrLf
    With rng
        .Font.Size = 11
        .Font.Bold = False
        .Font.Color = RGB(107, 114, 126)
        .Font.Italic = True
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .ParagraphFormat.SpaceAfter = 12
        .Collapse wdCollapseEnd
    End With
End Sub

' Add configuration form
Sub AddConfigForm(rng As Range)
    Dim tbl As Table
    
    ' Create 4x4 table
    Set tbl = ActiveDocument.Tables.Add(rng, 4, 4)
    Set ConfigTable = tbl
    
    With tbl
        .Style = "Plain Table 1"
        .Borders.Enable = False
        .Shading.BackgroundPatternColor = RGB(255, 255, 255)
        
        ' Column widths
        .Columns(1).Width = InchesToPoints(1.3)
        .Columns(2).Width = InchesToPoints(2.1)
        .Columns(3).Width = InchesToPoints(1.3)
        .Columns(4).Width = InchesToPoints(2.1)
        
        ' Row 1
        Call SetFormCell(.cell(1, 1), "GitLab URL", True)
        Call SetFormCell(.cell(1, 2), "https://gitlab.com", False)
        Call SetFormCell(.cell(1, 3), "Access Token", True)
        Call SetFormCell(.cell(1, 4), "glpat-xxxxxxxxxxxxxxxxxxxx", False)
        
        ' Row 2
        Call SetFormCell(.cell(2, 1), "Source Type", True)
        Call SetFormCell(.cell(2, 2), "Project", False)
        Call SetFormCell(.cell(2, 3), "Project/Group ID", True)
        Call SetFormCell(.cell(2, 4), "12345", False)
        
        ' Row 3
        Call SetFormCell(.cell(3, 1), "Target Type", True)
        Call SetFormCell(.cell(3, 2), "Epic", False)
        Call SetFormCell(.cell(3, 3), "Epic/Issue ID", True)
        Call SetFormCell(.cell(3, 4), "67", False)
        
        ' Row 4 - Status (merged)
        .cell(4, 1).Merge .cell(4, 4)
        Set StatusCell = .cell(4, 1).Range
        With StatusCell
            .text = ChrW(9679) & " Ready - Configure your GitLab connection above"
            .Font.Size = 10
            .Font.Color = RGB(34, 197, 94)
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Shading.BackgroundPatternColor = RGB(240, 253, 244)
            .ParagraphFormat.SpaceBefore = 6
            .ParagraphFormat.SpaceAfter = 6
        End With
        
        ' Table border
        With .Borders
            .OutsideLineStyle = wdLineStyleSingle
            .OutsideColor = RGB(229, 231, 235)
            .OutsideLineWidth = wdLineWidth075pt
        End With
    End With
    
    rng.SetRange tbl.Range.End + 1, tbl.Range.End + 1
    rng.text = vbCrLf
    rng.Collapse wdCollapseEnd
End Sub

' Set individual form cell
Sub SetFormCell(cell As cell, text As String, isLabel As Boolean)
    With cell.Range
        .text = text
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .ParagraphFormat.LeftIndent = InchesToPoints(0.08)
        
        If isLabel Then
            .Font.Bold = True
            .Font.Color = RGB(52, 53, 65)
        Else
            .Font.Bold = False
            .Font.Color = RGB(59, 130, 246)
            .Shading.BackgroundPatternColor = RGB(248, 250, 252)
        End If
    End With
End Sub

' Add action area with buttons
Sub AddActionArea(rng As Range)
    ' Button section header
    rng.text = ChrW(9733) & " Actions" & vbCrLf
    With rng
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(52, 53, 65)
        .ParagraphFormat.SpaceAfter = 6
        .Collapse wdCollapseEnd
    End With
    
    ' Create button table
    Dim btnTable As Table
    Set btnTable = ActiveDocument.Tables.Add(rng, 1, 3)
    
    With btnTable
        .Borders.Enable = False
        .Columns(1).Width = InchesToPoints(2.3)
        .Columns(2).Width = InchesToPoints(2.3)
        .Columns(3).Width = InchesToPoints(2.3)
        
        ' Button 1
        Call MakeButton(btnTable.cell(1, 1), ChrW(9654) & " EXTRACT", RGB(34, 197, 94))
        
        ' Button 2
        Call MakeButton(btnTable.cell(1, 2), ChrW(9881) & " COPILOT", RGB(139, 92, 246))
        
        ' Button 3
        Call MakeButton(btnTable.cell(1, 3), ChrW(10022) & " FORMAT", RGB(236, 72, 153))
    End With
    
    rng.SetRange btnTable.Range.End + 1, btnTable.Range.End + 1
    rng.text = vbCrLf
    rng.Collapse wdCollapseEnd
End Sub

' Make individual button
Sub MakeButton(cell As cell, buttonText As String, bgColor As Long)
    With cell.Range
        .text = buttonText
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Shading.BackgroundPatternColor = bgColor
        .ParagraphFormat.SpaceBefore = 8
        .ParagraphFormat.SpaceAfter = 8
    End With
    
    With cell.Borders
        .OutsideLineStyle = wdLineStyleSingle
        .OutsideColor = bgColor
'        .OutsideLineWidth = wdLineWidth150pt
    End With
End Sub

' Add instructions section
Sub AddInstructions(rng As Range)
    ' Shortcuts
    rng.text = ChrW(8250) & " Keyboard Shortcuts:" & vbCrLf
    With rng
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(52, 53, 65)
        .Collapse wdCollapseEnd
    End With
    
    rng.text = "Ctrl+Shift+G (Extract) " & ChrW(8226) & " Ctrl+Shift+C (Copilot) " & ChrW(8226) & " Ctrl+Shift+F (Format)" & vbCrLf & vbCrLf
    With rng
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(107, 114, 126)
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Collapse wdCollapseEnd
    End With
    
    ' Macro instructions
    rng.text = ChrW(9881) & " Alternative: Press Alt+F8 and run:" & vbCrLf
    With rng
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(52, 53, 65)
        .Collapse wdCollapseEnd
    End With
    
    rng.text = "RunExtract " & ChrW(8226) & " RunCopilot " & ChrW(8226) & " RunFormat " & ChrW(8226) & " TestConnection " & ChrW(8226) & " ExportCSV" & vbCrLf & vbCrLf
    With rng
        .Font.Size = 10
        .Font.Color = RGB(107, 114, 126)
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Collapse wdCollapseEnd
    End With
    
    ' Workflow
    rng.text = ChrW(8594) & " Workflow: Fill form " & ChrW(8594) & " EXTRACT " & ChrW(8594) & " COPILOT " & ChrW(8594) & " Copy results " & ChrW(8594) & " FORMAT"
    With rng
        .Font.Size = 10
        .Font.Color = RGB(142, 142, 160)
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Collapse wdCollapseEnd
    End With
End Sub

' === STATUS MANAGEMENT ===
Sub SetStatus(message As String, statusType As String)
    If StatusCell Is Nothing Then Exit Sub
    
    Dim icon As String
    Dim bgColor As Long
    Dim textColor As Long
    
    Select Case statusType
        Case "ready"
            icon = ChrW(9679)
            bgColor = RGB(240, 253, 244)
            textColor = RGB(34, 197, 94)
        Case "working"
            icon = ChrW(9881)
            bgColor = RGB(255, 251, 235)
            textColor = RGB(245, 158, 11)
        Case "success"
            icon = ChrW(9733)
            bgColor = RGB(240, 253, 244)
            textColor = RGB(34, 197, 94)
        Case "error"
            icon = ChrW(10060)
            bgColor = RGB(254, 242, 242)
            textColor = RGB(239, 68, 68)
    End Select
    
    With StatusCell
        .text = icon & " " & message
        .Font.Color = textColor
        .Shading.BackgroundPatternColor = bgColor
    End With
    
    Application.StatusBar = message
    DoEvents
End Sub

' === MAIN FUNCTIONS ===

' Extract GitLab data
Sub ExtractData()
    Dim config As Object
    Set config = ReadConfig()
    
    If config Is Nothing Then Exit Sub
    
    Call SetStatus("Connecting to GitLab...", "working")
    Call ClearResults
    
    If config("targetType") = "Epic" Then
        Call GetEpic(config)
    Else
        Call GetIssue(config)
    End If
End Sub

' Read configuration
Function ReadConfig() As Object
    Dim config As Object
    Set config = CreateObject("Scripting.Dictionary")
    
    If ConfigTable Is Nothing Then
        MsgBox "Please run StartGitLabGenerator first.", vbCritical
        Set ReadConfig = Nothing
        Exit Function
    End If
    
    On Error GoTo ConfigError
    
    With ConfigTable
        config("gitlabUrl") = CleanText(.cell(1, 2).Range.text)
        config("token") = CleanText(.cell(1, 4).Range.text)
        config("sourceType") = CleanText(.cell(2, 2).Range.text)
        config("sourceId") = CleanText(.cell(2, 4).Range.text)
        config("targetType") = CleanText(.cell(3, 2).Range.text)
        config("targetId") = CleanText(.cell(3, 4).Range.text)
    End With
    
    If Not ValidateConfig(config) Then
        Set ReadConfig = Nothing
        Exit Function
    End If
    
    Set ReadConfig = config
    Exit Function
    
ConfigError:
    Call SetStatus("Configuration error: " & Err.description, "error")
    Set ReadConfig = Nothing
End Function

' Clean text helper
Function CleanText(text As String) As String
    CleanText = Replace(Trim(text), Chr(13) & Chr(7), "")
End Function

' Validate configuration
Function ValidateConfig(config As Object) As Boolean
    ValidateConfig = False
    
    If config("gitlabUrl") = "" Or config("gitlabUrl") = "https://gitlab.com" Then
        Call SetStatus("Enter a valid GitLab URL", "error")
        Exit Function
    End If
    
    If config("token") = "" Or Left(config("token"), 5) = "glpat" And Len(config("token")) < 20 Then
        Call SetStatus("Enter a valid access token", "error")
        Exit Function
    End If
    
    If config("sourceId") = "" Or config("sourceId") = "12345" Then
        Call SetStatus("Enter a valid Project/Group ID", "error")
        Exit Function
    End If
    
    If config("targetId") = "" Or config("targetId") = "67" Then
        Call SetStatus("Enter a valid Epic/Issue ID", "error")
        Exit Function
    End If
    
    ValidateConfig = True
End Function

' Get Epic data
Sub GetEpic(config As Object)
    Dim apiUrl As String
    Dim epicData As String
    
    Call SetStatus("Fetching epic details...", "working")
    
    If InStr(UCase(config("sourceType")), "GROUP") > 0 Then
        apiUrl = config("gitlabUrl") & "/api/v4/groups/" & config("sourceId") & "/epics/" & config("targetId")
    Else
        apiUrl = config("gitlabUrl") & "/api/v4/projects/" & config("sourceId") & "/epics/" & config("targetId")
    End If
    
    epicData = CallAPI(apiUrl, config("token"))
    
    If epicData = "" Then
        Call SetStatus("Failed to fetch epic data", "error")
        Exit Sub
    End If
    
    Call ShowEpic(epicData)
    Call GetEpicIssues(config, config("targetId"))
    Call AddGenerationSection
    
    Call SetStatus("Epic loaded! Use COPILOT to generate test cases", "success")
End Sub

' Get single Issue data
Sub GetIssue(config As Object)
    Dim apiUrl As String
    Dim issueData As String
    
    Call SetStatus("Fetching issue details...", "working")
    
    apiUrl = config("gitlabUrl") & "/api/v4/projects/" & config("sourceId") & "/issues/" & config("targetId")
    issueData = CallAPI(apiUrl, config("token"))
    
    If issueData = "" Then
        Call SetStatus("Failed to fetch issue data", "error")
        Exit Sub
    End If
    
    Call ShowIssue(issueData, True)
    Call AddGenerationSection
    
    Call SetStatus("Issue loaded! Use COPILOT to generate test cases", "success")
End Sub

' Call GitLab API
Function CallAPI(url As String, token As String) As String
    Dim http As Object
    
    On Error GoTo APIError
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "Bearer " & token
    http.setRequestHeader "Content-Type", "application/json"
    http.send
    
    If http.Status = 200 Then
        CallAPI = http.responseText
    Else
        CallAPI = ""
    End If
    
    Exit Function
    
APIError:
    CallAPI = ""
End Function

' === COMPACT DATA DISPLAY ===

' Show epic in compact format
Sub ShowEpic(epicData As String)
    Dim doc As Document
    Dim rng As Range
    
    Set doc = ActiveDocument
    Set rng = doc.Range
    rng.Collapse wdCollapseEnd
    
    Dim title As String, description As String, state As String
    title = GetJSONValue(epicData, "title")
    description = GetJSONValue(epicData, "description")
    state = GetJSONValue(epicData, "state")
    
    ' Compact epic display
    Call AddDataSeparator(rng)
    
    rng.text = ChrW(9658) & " EPIC: " & title & vbCrLf
    With rng
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(52, 53, 65)
        .Collapse wdCollapseEnd
    End With
    
    rng.text = "Status: " & state & " " & ChrW(8226) & " " & Left(description, 100) & "..." & vbCrLf & vbCrLf
    With rng
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(107, 114, 126)
        .Collapse wdCollapseEnd
    End With
End Sub

' Show issue in compact format
Sub ShowIssue(issueData As String, isStandalone As Boolean)
    Dim doc As Document
    Dim rng As Range
    
    Set doc = ActiveDocument
    Set rng = doc.Range
    rng.Collapse wdCollapseEnd
    
    Dim title As String, description As String, state As String
    title = GetJSONValue(issueData, "title")
    description = GetJSONValue(issueData, "description")
    state = GetJSONValue(issueData, "state")
    
    If isStandalone Then
        Call AddDataSeparator(rng)
        rng.text = ChrW(9679) & " ISSUE: " & title & vbCrLf
        With rng
            .Font.Size = 14
            .Font.Bold = True
            .Font.Color = RGB(139, 92, 246)
        End With
    Else
        rng.text = ChrW(8250) & " " & title & vbCrLf
        With rng
            .Font.Size = 11
            .Font.Bold = True
            .Font.Color = RGB(139, 92, 246)
        End With
    End If
    
    With rng
        .Collapse wdCollapseEnd
        .text = "Status: " & state & " " & ChrW(8226) & " " & Left(description, 80) & "..." & vbCrLf & vbCrLf
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = RGB(107, 114, 126)
        .Collapse wdCollapseEnd
    End With
End Sub

' Get epic issues
Sub GetEpicIssues(config As Object, epicId As String)
    Dim apiUrl As String
    Dim issuesData As String
    
    Call SetStatus("Fetching linked issues...", "working")
    
    If InStr(UCase(config("sourceType")), "GROUP") > 0 Then
        apiUrl = config("gitlabUrl") & "/api/v4/groups/" & config("sourceId") & "/epics/" & epicId & "/issues"
    Else
        apiUrl = config("gitlabUrl") & "/api/v4/projects/" & config("sourceId") & "/epics/" & epicId & "/issues"
    End If
    
    issuesData = CallAPI(apiUrl, config("token"))
    
    If issuesData <> "" And issuesData <> "[]" Then
        Call ShowIssuesList(issuesData)
    End If
End Sub

' Show issues list in compact format
Sub ShowIssuesList(issuesData As String)
    Dim doc As Document
    Dim rng As Range
    Dim issueCount As Integer
    
    Set doc = ActiveDocument
    Set rng = doc.Range
    rng.Collapse wdCollapseEnd
    
    ' Count issues
    issueCount = CountOccurrences(issuesData, """title"":")
    
    rng.text = ChrW(8594) & " LINKED ISSUES (" & issueCount & " found)" & vbCrLf
    With rng
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(236, 72, 153)
        .Collapse wdCollapseEnd
    End With
    
    rng.text = "Complete details extracted for comprehensive test case generation." & vbCrLf & vbCrLf
    With rng
        .Font.Size = 9
        .Font.Color = RGB(107, 114, 126)
        .Collapse wdCollapseEnd
    End With
End Sub

' Add data separator
Sub AddDataSeparator(rng As Range)
    rng.text = vbCrLf & String(50, ChrW(9552)) & vbCrLf
    With rng
        .Font.Color = RGB(229, 231, 235)
        .Collapse wdCollapseEnd
    End With
End Sub

' Add generation section
Sub AddGenerationSection()
    Dim doc As Document
    Dim rng As Range
    
    Set doc = ActiveDocument
    Set rng = doc.Range
    rng.Collapse wdCollapseEnd
    
    Call AddDataSeparator(rng)
    
    rng.text = ChrW(9881) & " AI TEST CASE GENERATION" & vbCrLf
    With rng
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = RGB(139, 92, 246)
        .Collapse wdCollapseEnd
    End With
    
    rng.text = "Ready to generate comprehensive test cases for the data above." & vbCrLf & vbCrLf
    With rng
        .Font.Size = 10
        .Font.Color = RGB(52, 53, 65)
        .Collapse wdCollapseEnd
    End With
    
    Call CreateCompactTable(rng)
End Sub

' Create compact test case table
Sub CreateCompactTable(rng As Range)
    Dim tbl As Table
    
    Set tbl = ActiveDocument.Tables.Add(rng, 3, 6)
    
    With tbl
        ' Headers
        .cell(1, 1).Range.text = "Test ID"
        .cell(1, 2).Range.text = "Category"
        .cell(1, 3).Range.text = "Scenario"
        .cell(1, 4).Range.text = "Steps"
        .cell(1, 5).Range.text = "Expected"
        .cell(1, 6).Range.text = "Priority"
        
        ' Header styling
        With .Rows(1).Range
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Shading.BackgroundPatternColor = RGB(52, 53, 65)
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
        End With
        
        ' Compact column widths
        .Columns(1).Width = InchesToPoints(0.7)
        .Columns(2).Width = InchesToPoints(1#)
        .Columns(3).Width = InchesToPoints(1.8)
        .Columns(4).Width = InchesToPoints(2.2)
        .Columns(5).Width = InchesToPoints(1.5)
        .Columns(6).Width = InchesToPoints(0.8)
        
        ' Clean styling
        .Style = "Plain Table 1"
        .Borders.OutsideLineStyle = wdLineStyleSingle
        .Borders.OutsideColor = RGB(229, 231, 235)
        .Borders.InsideLineStyle = wdLineStyleSingle
        .Borders.InsideColor = RGB(243, 244, 246)
    End With
    
    rng.SetRange tbl.Range.End + 1, tbl.Range.End + 1
    rng.text = vbCrLf & ChrW(10022) & " Use Copilot to generate, paste results, then FORMAT" & vbCrLf
    With rng
        .Font.Size = 9
        .Font.Color = RGB(142, 142, 160)
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Collapse wdCollapseEnd
    End With
End Sub

' === COPILOT INTEGRATION ===

' Trigger Copilot
Sub TriggerCopilot()
    Call SetStatus("Opening Copilot...", "working")
    
    If OpenCopilot() Then
        Call SetStatus("Copilot activated!", "success")
    Else
        Call SetStatus("Use manual Copilot activation", "working")
        Call ShowCopilotHelp
    End If
End Sub

' Try to open Copilot
Function OpenCopilot() As Boolean
    On Error GoTo Manual
    Application.CommandBars.ExecuteMso "CopilotToggle"
    OpenCopilot = True
    Exit Function
    
Manual:
    OpenCopilot = False
End Function

' Show Copilot help
Sub ShowCopilotHelp()
    MsgBox ChrW(9881) & " Manual Copilot Steps:" & vbCrLf & vbCrLf & _
           "1. Click Copilot button in Word ribbon" & vbCrLf & _
           "2. Type: 'Generate test cases for this document'" & vbCrLf & _
           "3. Copy generated test cases" & vbCrLf & _
           "4. Paste in table above" & vbCrLf & _
           "5. Run FORMAT to clean up", vbInformation
End Sub

' === FORMATTING ===

' Format test cases
Sub FormatResults()
    Dim doc As Document
    Dim tbl As Table
    
    Set doc = ActiveDocument
    
    If doc.Tables.count = 0 Then
        Call SetStatus("No table found to format", "error")
        Exit Sub
    End If
    
    Set tbl = doc.Tables(doc.Tables.count)
    
    Call SetStatus("Formatting test cases...", "working")
    Call CleanTable(tbl)
    Call SetStatus("Formatting complete!", "success")
End Sub

' Clean and format table
Sub CleanTable(tbl As Table)
    Dim i As Integer, j As Integer
    
    ' Clean content and apply formatting
    For i = 2 To tbl.Rows.count
        For j = 1 To tbl.Columns.count
            On Error Resume Next
            Dim cellText As String
            cellText = tbl.cell(i, j).Range.text
            cellText = Replace(cellText, Chr(13) & Chr(7), "")
            cellText = Replace(cellText, "**", "")
            cellText = Trim(cellText)
            tbl.cell(i, j).Range.text = cellText
            
            With tbl.cell(i, j).Range
                .Font.Name = "Segoe UI"
                .Font.Size = 9
                
                Select Case j
                    Case 1 ' Test ID
                        .Font.Bold = True
                        .Font.Color = RGB(59, 130, 246)
                        .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    Case 2 ' Category
                        .Font.Color = RGB(34, 197, 94)
                        .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    Case 6 ' Priority
                        .Font.Bold = True
                        .ParagraphFormat.Alignment = wdAlignParagraphCenter
                        Call ColorPriority(tbl.cell(i, j))
                    Case Else
                        .Font.Color = RGB(52, 53, 65)
                        .ParagraphFormat.Alignment = wdAlignParagraphLeft
                End Select
            End With
            On Error GoTo 0
        Next j
    Next i
    
    ' Final table styling
    With tbl
        .AutoFitBehavior wdAutoFitContent
        .Borders.Enable = True
    End With
End Sub

' Color priority cells
Sub ColorPriority(cell As cell)
    Dim text As String
    text = UCase(cell.Range.text)
    
    With cell.Range
        If InStr(text, "HIGH") > 0 Or InStr(text, "CRITICAL") > 0 Then
            .Font.Color = RGB(255, 255, 255)
            .Shading.BackgroundPatternColor = RGB(239, 68, 68)
        ElseIf InStr(text, "MEDIUM") > 0 Then
            .Font.Color = RGB(0, 0, 0)
            .Shading.BackgroundPatternColor = RGB(245, 158, 11)
        ElseIf InStr(text, "LOW") > 0 Then
            .Font.Color = RGB(255, 255, 255)
            .Shading.BackgroundPatternColor = RGB(34, 197, 94)
        End If
    End With
End Sub

' === UTILITY FUNCTIONS ===

' Clear results area
Sub ClearResults()
    Dim doc As Document
    Dim rng As Range
    
    Set doc = ActiveDocument
    
    ' Keep only first 2 tables (config and buttons)
    If doc.Tables.count > 2 Then
        Set rng = doc.Range(doc.Tables(2).Range.End, doc.Range.End)
        rng.Delete
    End If
End Sub

' Clear entire document
Sub ClearDocument()
    ActiveDocument.Range.Delete
    Set ConfigTable = Nothing
    Set StatusCell = Nothing
End Sub

' Setup keyboard shortcuts
Sub SetupShortcuts()
    On Error Resume Next
    CustomizationContext = ActiveDocument
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyG), _
                   KeyCategory:=wdKeyCategoryMacro, _
                   Command:="RunExtract"
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyC), _
                   KeyCategory:=wdKeyCategoryMacro, _
                   Command:="RunCopilot"
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyF), _
                   KeyCategory:=wdKeyCategoryMacro, _
                   Command:="RunFormat"
    
    On Error GoTo 0
End Sub

' Get JSON value (simple parser)
Function GetJSONValue(json As String, key As String) As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim searchStr As String
    
    GetJSONValue = "Not available"
    
    searchStr = """" & key & """:"
    startPos = InStr(json, searchStr)
    
    If startPos > 0 Then
        startPos = startPos + Len(searchStr)
        
        ' Skip whitespace and quotes
        Do While Mid(json, startPos, 1) = " " Or Mid(json, startPos, 1) = """"
            startPos = startPos + 1
        Loop
        
        ' Find end
        endPos = startPos
        Do While endPos <= Len(json)
            Dim char As String
            char = Mid(json, endPos, 1)
            If char = """" Or char = "," Or char = "}" Then Exit Do
            endPos = endPos + 1
        Loop
        
        If endPos > startPos Then
            GetJSONValue = Mid(json, startPos, endPos - startPos)
            GetJSONValue = Replace(GetJSONValue, "\n", " ")
            GetJSONValue = Trim(GetJSONValue)
        End If
    End If
End Function

' Count occurrences helper
Function CountOccurrences(text As String, searchString As String) As Integer
    Dim count As Integer
    Dim pos As Integer
    
    count = 0
    pos = 1
    
    Do
        pos = InStr(pos, text, searchString)
        If pos > 0 Then
            count = count + 1
            pos = pos + Len(searchString)
        End If
    Loop While pos > 0
    
    CountOccurrences = count
End Function

' === MACRO RUNNERS (for Alt+F8) ===

Sub RunExtract()
    ExtractData
End Sub

Sub RunCopilot()
    TriggerCopilot
End Sub

Sub RunFormat()
    FormatResults
End Sub

Sub TestConnection()
    Dim config As Object
    Set config = ReadConfig()
    
    If config Is Nothing Then Exit Sub
    
    Call SetStatus("Testing connection...", "working")
    
    Dim testUrl As String
    testUrl = config("gitlabUrl") & "/api/v4/user"
    
    Dim response As String
    response = CallAPI(testUrl, config("token"))
    
    If response <> "" Then
        Call SetStatus("Connection successful!", "success")
        MsgBox ChrW(9733) & " GitLab connection works!", vbInformation
    Else
        Call SetStatus("Connection failed", "error")
        MsgBox ChrW(10060) & " Check your URL and token", vbCritical
    End If
End Sub

Sub ExportCSV()
    Dim doc As Document
    Dim tbl As Table
    Dim csvContent As String
    Dim fileName As String
    Dim i As Integer, j As Integer
    
    Set doc = ActiveDocument
    
    If doc.Tables.count = 0 Then
        MsgBox "No table to export", vbExclamation
        Exit Sub
    End If
    
    Set tbl = doc.Tables(doc.Tables.count)
    
    ' Build CSV
    For i = 1 To tbl.Rows.count
        For j = 1 To tbl.Columns.count
            On Error Resume Next
            Dim cellText As String
            cellText = tbl.cell(i, j).Range.text
            cellText = Replace(cellText, Chr(13) & Chr(7), "")
            cellText = Replace(cellText, """", """""")
            csvContent = csvContent & """" & cellText & """"
            If j < tbl.Columns.count Then csvContent = csvContent & ","
            On Error GoTo 0
        Next j
        csvContent = csvContent & vbCrLf
    Next i
    
    ' Save file
    fileName = "TestCases_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"
    Dim filePath As String
    filePath = Environ("USERPROFILE") & "\Desktop\" & fileName
    
    Open filePath For Output As #1
    Print #1, csvContent
    Close #1
    
    Call SetStatus("Exported to " & fileName, "success")
    MsgBox ChrW(9733) & " Exported to Desktop\" & fileName, vbInformation
End Sub

Sub QuickSetup()
    If ConfigTable Is Nothing Then
        MsgBox "Run StartGitLabGenerator first", vbExclamation
        Exit Sub
    End If
    
    ConfigTable.cell(1, 2).Range.text = "https://gitlab.com"
    ConfigTable.cell(2, 2).Range.text = "Project"
    ConfigTable.cell(3, 2).Range.text = "Epic"
    
    Call SetStatus("Pre-configured for GitLab.com", "ready")
End Sub


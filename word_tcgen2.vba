' GitLab Test Case Generator - Enhanced with Detailed Descriptions
' Now includes full descriptions, page breaks, and comprehensive data extraction

Option Explicit

' Global variables
Dim ConfigTable As Table
Dim StatusCell As Range
Dim FirstPageEnd As Long

' === MAIN ENTRY POINT ===
Sub StartGitLabGenerator()
    Call ClearDocument
    Call CreateInterface
    Call SetupShortcuts
    
    ' Mark end of first page for protection
    FirstPageEnd = ActiveDocument.Range.End
    
    MsgBox ChrW(9733) & " GitLab Test Case Generator Ready!" & vbCrLf & vbCrLf & _
           "Use Alt+F8 to run macros or keyboard shortcuts!" & vbCrLf & _
           "Enhanced with detailed descriptions and page protection!", vbInformation
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
    Call AddCopilotPrompt(rng)
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
    rng.text = "AI-powered test case generation with detailed GitLab integration" & vbCrLf & vbCrLf
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
    rng.text = ChrW(8594) & " Enhanced Workflow: Fill form " & ChrW(8594) & " EXTRACT (detailed data) " & ChrW(8594) & " COPILOT (use prompt below) " & ChrW(8594) & " FORMAT"
    With rng
        .Font.Size = 10
        .Font.Color = RGB(142, 142, 160)
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Collapse wdCollapseEnd
    End With
End Sub

' Add comprehensive Copilot prompt
Sub AddCopilotPrompt(rng As Range)
    rng.text = vbCrLf & ChrW(9881) & " COMPREHENSIVE COPILOT PROMPT" & vbCrLf
    With rng
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(139, 92, 246)
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Collapse wdCollapseEnd
    End With
    
    rng.text = "Copy this prompt to M365 Copilot after extracting GitLab data:" & vbCrLf & vbCrLf
    With rng
        .Font.Size = 10
        .Font.Color = RGB(107, 114, 126)
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Collapse wdCollapseEnd
    End With
    
    ' Create prompt box
    Dim promptTable As Table
    Set promptTable = ActiveDocument.Tables.Add(rng, 1, 1)
    
    With promptTable
        .cell(1, 1).Range.text = GetComprehensiveCopilotPrompt()
        With .cell(1, 1).Range
            .Font.Name = "Consolas"
            .Font.Size = 9
            .Font.Color = RGB(52, 53, 65)
            .Shading.BackgroundPatternColor = RGB(248, 250, 252)
            .ParagraphFormat.LeftIndent = InchesToPoints(0.1)
            .ParagraphFormat.RightIndent = InchesToPoints(0.1)
        End With
        
        With .Borders
            .OutsideLineStyle = wdLineStyleSingle
            .OutsideColor = RGB(139, 92, 246)
            .OutsideLineWidth = wdLineWidth150pt
        End With
        
        .AutoFitBehavior wdAutoFitWindow
    End With
    
    rng.SetRange promptTable.Range.End + 1, promptTable.Range.End + 1
    rng.text = vbCrLf
    rng.Collapse wdCollapseEnd
End Sub

' Generate comprehensive Copilot prompt
Function GetComprehensiveCopilotPrompt() As String
    GetComprehensiveCopilotPrompt = _
"Act as a Senior QA Test Architect and analyze all GitLab epics, issues, and detailed descriptions in this document to generate comprehensive test cases." & vbCrLf & vbCrLf & _

"ANALYSIS REQUIREMENTS:" & vbCrLf & _
"• Extract ALL functional requirements from epic/issue descriptions" & vbCrLf & _
"• Identify user stories, acceptance criteria, and technical specifications" & vbCrLf & _
"• Note any mentioned edge cases, error conditions, or constraints" & vbCrLf & _
"• Consider integration points, dependencies, and workflows" & vbCrLf & _
"• Analyze security, performance, and accessibility requirements" & vbCrLf & vbCrLf & _

"TEST CASE GENERATION:" & vbCrLf & _
"Generate test cases covering:" & vbCrLf & _
"1. FUNCTIONAL TESTING - Core features and user journeys" & vbCrLf & _
"2. BOUNDARY TESTING - Input validation and limits" & vbCrLf & _
"3. NEGATIVE TESTING - Error handling and invalid scenarios" & vbCrLf & _
"4. INTEGRATION TESTING - System interactions and data flow" & vbCrLf & _
"5. UI/UX TESTING - Interface elements and user experience" & vbCrLf & _
"6. SECURITY TESTING - Authentication, authorization, data protection" & vbCrLf & _
"7. PERFORMANCE TESTING - Load, stress, and response time scenarios" & vbCrLf & _
"8. ACCESSIBILITY TESTING - Compliance and usability standards" & vbCrLf & vbCrLf & _

"OUTPUT FORMAT:" & vbCrLf & _
"Create a detailed table with these columns:" & vbCrLf & _
"• Test ID: TC-[Category]-[Number] (e.g., TC-FUNC-001)" & vbCrLf & _
"• Category: FUNCTIONAL, INTEGRATION, SECURITY, PERFORMANCE, UI, NEGATIVE, BOUNDARY, ACCESSIBILITY" & vbCrLf & _
"• Test Scenario: Clear, specific scenario description" & vbCrLf & _
"• Test Steps: Detailed step-by-step instructions (numbered)" & vbCrLf & _
"• Expected Result: Specific expected outcome" & vbCrLf & _
"• Priority: HIGH/MEDIUM/LOW based on business impact" & vbCrLf & vbCrLf & _

"QUALITY REQUIREMENTS:" & vbCrLf & _
"• Minimum 25-40 comprehensive test cases" & vbCrLf & _
"• Include both positive and negative test scenarios" & vbCrLf & _
"• Cover all user roles and permission levels mentioned" & vbCrLf & _
"• Address cross-browser/device compatibility if web-based" & vbCrLf & _
"• Include data validation and error message verification" & vbCrLf & _
"• Consider regression testing for existing functionality" & vbCrLf & vbCrLf & _

"SPECIAL CONSIDERATIONS:" & vbCrLf & _
"• Pay attention to acceptance criteria in issue descriptions" & vbCrLf & _
"• Extract test scenarios from user story descriptions" & vbCrLf & _
"• Consider API testing if backend services are mentioned" & vbCrLf & _
"• Include mobile responsiveness testing if applicable" & vbCrLf & _
"• Address compliance requirements (GDPR, WCAG, etc.)" & vbCrLf & vbCrLf & _

"Please analyze the complete GitLab data below and generate comprehensive test cases following this framework."
End Function

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

' Extract GitLab data with page break protection
Sub ExtractData()
    Dim config As Object
    Set config = ReadConfig()
    
    If config Is Nothing Then Exit Sub
    
    Call SetStatus("Connecting to GitLab...", "working")
    Call ClearResults
    
    ' Insert page break to protect first page
    Call InsertPageBreak
    
    If config("targetType") = "Epic" Then
        Call GetEpicDetailed(config)
    Else
        Call GetIssueDetailed(config)
    End If
End Sub

' Insert page break to separate config from data
Sub InsertPageBreak()
    Dim doc As Document
    Dim rng As Range
    
    Set doc = ActiveDocument
    Set rng = doc.Range
    rng.Collapse wdCollapseEnd
    
    rng.text = vbCrLf & vbCrLf
    rng.Collapse wdCollapseEnd
    rng.InsertBreak Type:=wdPageBreak
    
    rng.text = vbCrLf
    rng.Collapse wdCollapseEnd
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

' Get Epic data with full details
Sub GetEpicDetailed(config As Object)
    Dim apiUrl As String
    Dim epicData As String
    
    Call SetStatus("Fetching detailed epic information...", "working")
    
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
    
    Call ShowEpicDetailed(epicData)
    Call GetEpicIssuesDetailed(config, config("targetId"))
    Call AddGenerationSection
    
    Call SetStatus("Detailed epic data extracted! Ready for Copilot analysis", "success")
End Sub

' Get single Issue data with full details
Sub GetIssueDetailed(config As Object)
    Dim apiUrl As String
    Dim issueData As String
    
    Call SetStatus("Fetching detailed issue information...", "working")
    
    apiUrl = config("gitlabUrl") & "/api/v4/projects/" & config("sourceId") & "/issues/" & config("targetId")
    issueData = CallAPI(apiUrl, config("token"))
    
    If issueData = "" Then
        Call SetStatus("Failed to fetch issue data", "error")
        Exit Sub
    End If
    
    Call ShowIssueDetailed(issueData, True)
    Call AddGenerationSection
    
    Call SetStatus("Detailed issue data extracted! Ready for Copilot analysis", "success")
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

' === DETAILED DATA DISPLAY ===

' Show epic with full details
Sub ShowEpicDetailed(epicData As String)
    Dim doc As Document
    Dim rng As Range
    
    Set doc = ActiveDocument
    Set rng = doc.Range
    rng.Collapse wdCollapseEnd
    
    ' Extract detailed data
    Dim title As String, description As String, state As String
    Dim author As String, createdAt As String, updatedAt As String
    Dim labels As String, webUrl As String
    
    title = GetJSONValue(epicData, "title")
    description = GetJSONValueLong(epicData, "description")
    state = GetJSONValue(epicData, "state")
    author = GetJSONValue(epicData, "author")
    createdAt = GetJSONValue(epicData, "created_at")
    updatedAt = GetJSONValue(epicData, "updated_at")
    webUrl = GetJSONValue(epicData, "web_url")
    
    ' Main epic header
    Call AddDataSeparator(rng)
    
    rng.text = ChrW(9658) & " EPIC ANALYSIS: " & title & vbCrLf
    With rng
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(52, 53, 65)
        .Collapse wdCollapseEnd
    End With
    
    ' Epic metadata
    rng.text = "Status: " & state & " | Created: " & FormatDate(createdAt) & " | Updated: " & FormatDate(updatedAt) & vbCrLf
    With rng
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(107, 114, 126)
        .Collapse wdCollapseEnd
    End With
    
    rng.text = "URL: " & webUrl & vbCrLf & vbCrLf
    With rng
        .Font.Size = 9
        .Font.Color = RGB(59, 130, 246)
        .Collapse wdCollapseEnd
    End With
    
    ' Detailed description
    If Len(description) > 10 Then
        rng.text = ChrW(9733) & " DETAILED EPIC DESCRIPTION:" & vbCrLf
        With rng
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = RGB(34, 197, 94)
            .Collapse wdCollapseEnd
        End With
        
        Call ShowFormattedDescription(rng, description)
    End If
End Sub

' Show issue with full details
Sub ShowIssueDetailed(issueData As String, isStandalone As Boolean)
    Dim doc As Document
    Dim rng As Range
    
    Set doc = ActiveDocument
    Set rng = doc.Range
    rng.Collapse wdCollapseEnd
    
    ' Extract detailed data
    Dim title As String, description As String, state As String
    Dim author As String, createdAt As String, updatedAt As String
    Dim labels As String, webUrl As String, issueType As String
    
    title = GetJSONValue(issueData, "title")
    description = GetJSONValueLong(issueData, "description")
    state = GetJSONValue(issueData, "state")
    author = GetJSONValue(issueData, "author")
    createdAt = GetJSONValue(issueData, "created_at")
    updatedAt = GetJSONValue(issueData, "updated_at")
    webUrl = GetJSONValue(issueData, "web_url")
    issueType = GetJSONValue(issueData, "issue_type")
    
    If isStandalone Then
        Call AddDataSeparator(rng)
        rng.text = ChrW(9679) & " ISSUE ANALYSIS: " & title & vbCrLf
        With rng
            .Font.Size = 16
            .Font.Bold = True
            .Font.Color = RGB(139, 92, 246)
        End With
    Else
        rng.text = vbCrLf & ChrW(8250) & " LINKED ISSUE: " & title & vbCrLf
        With rng
            .Font.Size = 13
            .Font.Bold = True
            .Font.Color = RGB(139, 92, 246)
        End With
    End If
    
    With rng
        .Collapse wdCollapseEnd
    End With
    
    ' Issue metadata
    rng.text = "Status: " & state & " | Type: " & issueType & " | Created: " & FormatDate(createdAt) & vbCrLf
    With rng
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(107, 114, 126)
        .Collapse wdCollapseEnd
    End With
    
    rng.text = "URL: " & webUrl & vbCrLf & vbCrLf
    With rng
        .Font.Size = 9
        .Font.Color = RGB(59, 130, 246)
        .Collapse wdCollapseEnd
    End With
    
    ' Detailed description
    If Len(description) > 10 Then
        rng.text = ChrW(9733) & " DETAILED ISSUE DESCRIPTION:" & vbCrLf
        With rng
            .Font.Size = 12
            .Font.Bold = True
            .Font.Color = RGB(34, 197, 94)
            .Collapse wdCollapseEnd
        End With
        
        Call ShowFormattedDescription(rng, description)
    End If
End Sub

' Show formatted description with proper structure
Sub ShowFormattedDescription(rng As Range, description As String)
    ' Clean and format the description
    description = Replace(description, "\r\n", vbCrLf)
    description = Replace(description, "\n", vbCrLf)
    description = Replace(description, "\\", "\")
    description = Replace(description, "\""", """")
    
    rng.text = description & vbCrLf & vbCrLf
    With rng
        .Font.Size = 10
        .Font.Bold = False
        .Font.Color = RGB(52, 53, 65)
        .ParagraphFormat.LeftIndent = InchesToPoints(0.2)
        .Shading.BackgroundPatternColor = RGB(248, 250, 252)
        .Collapse wdCollapseEnd
    End With
End Sub

' Get epic issues with detailed descriptions
Sub GetEpicIssuesDetailed(config As Object, epicId As String)
    Dim apiUrl As String
    Dim issuesData As String
    
    Call SetStatus("Fetching detailed linked issues...", "working")
    
    If InStr(UCase(config("sourceType")), "GROUP") > 0 Then
        apiUrl = config("gitlabUrl") & "/api/v4/groups/" & config("sourceId") & "/epics/" & epicId & "/issues"
    Else
        apiUrl = config("gitlabUrl") & "/api/v4/projects/" & config("sourceId") & "/epics/" & epicId & "/issues"
    End If
    
    issuesData = CallAPI(apiUrl, config("token"))
    
    If issuesData <> "" And issuesData <> "[]" Then
        Call ShowDetailedIssuesList(issuesData, config)
    End If
End Sub

' Show detailed issues list with full descriptions
Sub ShowDetailedIssuesList(issuesData As String, config As Object)
    Dim doc As Document
    Dim rng As Range
    Dim issueCount As Integer
    
    Set doc = ActiveDocument
    Set rng = doc.Range
    rng.Collapse wdCollapseEnd
    
    ' Count issues
    issueCount = CountOccurrences(issuesData, """title"":")
    
    rng.text = ChrW(8594) & " DETAILED LINKED ISSUES (" & issueCount & " found)" & vbCrLf
    With rng
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(236, 72, 153)
        .Collapse wdCollapseEnd
    End With
    
    ' Extract and show each issue with details
    Call ProcessIssuesFromList(issuesData, config)
End Sub

' Process individual issues from the list
Sub ProcessIssuesFromList(issuesData As String, config As Object)
    Dim issueIds As Collection
    Set issueIds = ExtractIssueIds(issuesData)
    
    Dim i As Integer
    For i = 1 To issueIds.count
        If i <= 10 Then ' Limit to prevent excessive API calls
            Call GetAndShowIndividualIssue(config, issueIds(i))
        End If
    Next i
    
    If issueIds.count > 10 Then
        Dim doc As Document
        Dim rng As Range
        Set doc = ActiveDocument
        Set rng = doc.Range
        rng.Collapse wdCollapseEnd
        
        rng.text = vbCrLf & "Note: Showing first 10 issues. Total found: " & issueIds.count & vbCrLf & vbCrLf
        With rng
            .Font.Size = 9
            .Font.Color = RGB(107, 114, 126)
            .Font.Italic = True
            .Collapse wdCollapseEnd
        End With
    End If
End Sub

' Extract issue IDs from JSON list
Function ExtractIssueIds(issuesData As String) As Collection
    Dim ids As New Collection
    Dim pos As Integer
    Dim searchStr As String
    
    searchStr = """iid"":"
    pos = 1
    
    Do
        pos = InStr(pos, issuesData, searchStr)
        If pos > 0 Then
            pos = pos + Len(searchStr)
            ' Skip whitespace
            Do While Mid(issuesData, pos, 1) = " "
                pos = pos + 1
            Loop
            
            ' Extract ID number
            Dim idStart As Integer
            Dim idEnd As Integer
            idStart = pos
            idEnd = pos
            
            Do While IsNumeric(Mid(issuesData, idEnd, 1)) And idEnd <= Len(issuesData)
                idEnd = idEnd + 1
            Loop
            
            If idEnd > idStart Then
                ids.Add Mid(issuesData, idStart, idEnd - idStart)
            End If
            
            pos = idEnd
        End If
    Loop While pos > 0 And pos < Len(issuesData)
    
    Set ExtractIssueIds = ids
End Function

' Get and show individual issue with full details
Sub GetAndShowIndividualIssue(config As Object, issueId As String)
    Dim apiUrl As String
    Dim issueData As String
    
    apiUrl = config("gitlabUrl") & "/api/v4/projects/" & config("sourceId") & "/issues/" & issueId
    issueData = CallAPI(apiUrl, config("token"))
    
    If issueData <> "" Then
        Call ShowIssueDetailed(issueData, False)
    End If
End Sub

' Add data separator
Sub AddDataSeparator(rng As Range)
    rng.text = vbCrLf & String(80, ChrW(9552)) & vbCrLf
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
    
    rng.text = ChrW(9881) & " AI TEST CASE GENERATION READY" & vbCrLf
    With rng
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = RGB(139, 92, 246)
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Collapse wdCollapseEnd
    End With
    
    rng.text = "All GitLab data extracted with detailed descriptions." & vbCrLf & _
               "Use COPILOT button or copy the comprehensive prompt from page 1." & vbCrLf & vbCrLf
    With rng
        .Font.Size = 11
        .Font.Color = RGB(52, 53, 65)
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Collapse wdCollapseEnd
    End With
    
    Call CreateComprehensiveTestTable(rng)
End Sub

' Create comprehensive test case table
Sub CreateComprehensiveTestTable(rng As Range)
    Dim tbl As Table
    
    Set tbl = ActiveDocument.Tables.Add(rng, 5, 6)
    
    With tbl
        ' Headers
        .cell(1, 1).Range.text = "Test ID"
        .cell(1, 2).Range.text = "Category"
        .cell(1, 3).Range.text = "Test Scenario"
        .cell(1, 4).Range.text = "Test Steps"
        .cell(1, 5).Range.text = "Expected Result"
        .cell(1, 6).Range.text = "Priority"
        
        ' Sample rows for guidance
        .cell(2, 1).Range.text = "TC-FUNC-001"
        .cell(2, 2).Range.text = "FUNCTIONAL"
        .cell(2, 3).Range.text = "[Copilot will generate scenarios]"
        .cell(2, 4).Range.text = "[Detailed test steps]"
        .cell(2, 5).Range.text = "[Expected outcomes]"
        .cell(2, 6).Range.text = "HIGH"
        
        .cell(3, 1).Range.text = "TC-INT-001"
        .cell(3, 2).Range.text = "INTEGRATION"
        .cell(3, 3).Range.text = "[Integration test scenarios]"
        .cell(3, 4).Range.text = "[API/System integration steps]"
        .cell(3, 5).Range.text = "[Integration results]"
        .cell(3, 6).Range.text = "MEDIUM"
        
        .cell(4, 1).Range.text = "TC-SEC-001"
        .cell(4, 2).Range.text = "SECURITY"
        .cell(4, 3).Range.text = "[Security test scenarios]"
        .cell(4, 4).Range.text = "[Security validation steps]"
        .cell(4, 5).Range.text = "[Security compliance]"
        .cell(4, 6).Range.text = "HIGH"
        
        .cell(5, 1).Range.text = "TC-NEG-001"
        .cell(5, 2).Range.text = "NEGATIVE"
        .cell(5, 3).Range.text = "[Error handling scenarios]"
        .cell(5, 4).Range.text = "[Invalid input testing]"
        .cell(5, 5).Range.text = "[Proper error messages]"
        .cell(5, 6).Range.text = "MEDIUM"
        
        ' Header styling
        With .Rows(1).Range
            .Font.Name = "Segoe UI"
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Shading.BackgroundPatternColor = RGB(52, 53, 65)
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
        End With
        
        ' Sample row styling
        With .Range(2, 1, 5, 6)
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .Font.Color = RGB(107, 114, 126)
            .Font.Italic = True
        End With
        
        ' Optimized column widths
        .Columns(1).Width = InchesToPoints(0.9)   ' Test ID
        .Columns(2).Width = InchesToPoints(1.1)   ' Category
        .Columns(3).Width = InchesToPoints(2.2)   ' Scenario
        .Columns(4).Width = InchesToPoints(2.5)   ' Steps
        .Columns(5).Width = InchesToPoints(2.0)   ' Expected
        .Columns(6).Width = InchesToPoints(0.8)   ' Priority
        
        ' Professional styling
        .Style = "Plain Table 1"
        .Borders.OutsideLineStyle = wdLineStyleSingle
        .Borders.OutsideColor = RGB(52, 53, 65)
        .Borders.OutsideLineWidth = wdLineWidth150pt
        .Borders.InsideLineStyle = wdLineStyleSingle
        .Borders.InsideColor = RGB(229, 231, 235)
        .Borders.InsideLineWidth = wdLineWidth075pt
    End With
    
    rng.SetRange tbl.Range.End + 1, tbl.Range.End + 1
    rng.text = vbCrLf & ChrW(10022) & " INSTRUCTIONS:" & vbCrLf & _
               "1. Use Copilot with the comprehensive prompt from page 1" & vbCrLf & _
               "2. Replace sample rows with generated test cases" & vbCrLf & _
               "3. Run FORMAT macro to apply professional styling" & vbCrLf & _
               "4. Export to CSV if needed for test management tools" & vbCrLf
    With rng
        .Font.Size = 10
        .Font.Color = RGB(52, 53, 65)
        .Font.Bold = True
        .ParagraphFormat.LeftIndent = InchesToPoints(0.2)
        .Collapse wdCollapseEnd
    End With
End Sub

' === COPILOT INTEGRATION ===

' Trigger Copilot
Sub TriggerCopilot()
    Call SetStatus("Opening Copilot for comprehensive analysis...", "working")
    
    If OpenCopilot() Then
        Call SetStatus("Copilot activated! Use the comprehensive prompt.", "success")
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
    MsgBox ChrW(9881) & " Enhanced Copilot Workflow:" & vbCrLf & vbCrLf & _
           "1. Click Copilot button in Word ribbon" & vbCrLf & _
           "2. Copy the comprehensive prompt from page 1" & vbCrLf & _
           "3. Paste into Copilot chat" & vbCrLf & _
           "4. Let Copilot analyze all detailed descriptions" & vbCrLf & _
           "5. Copy generated test cases into the table" & vbCrLf & _
           "6. Run FORMAT for professional styling" & vbCrLf & vbCrLf & _
           "The enhanced prompt covers functional, security," & vbCrLf & _
           "integration, and negative testing scenarios!", vbInformation
End Sub

' === FORMATTING ===

' Format test cases with enhanced styling
Sub FormatResults()
    Dim doc As Document
    Dim tbl As Table
    
    Set doc = ActiveDocument
    
    If doc.Tables.count = 0 Then
        Call SetStatus("No table found to format", "error")
        Exit Sub
    End If
    
    Set tbl = doc.Tables(doc.Tables.count)
    
    Call SetStatus("Applying comprehensive formatting...", "working")
    Call CleanAndStyleTable(tbl)
    Call SetStatus("Professional formatting complete!", "success")
End Sub

' Clean and format table with enhanced styling
Sub CleanAndStyleTable(tbl As Table)
    Dim i As Integer, j As Integer
    
    ' Clean content and apply enhanced formatting
    For i = 2 To tbl.Rows.count
        For j = 1 To tbl.Columns.count
            On Error Resume Next
            Dim cellText As String
            cellText = tbl.cell(i, j).Range.text
            cellText = Replace(cellText, Chr(13) & Chr(7), "")
            cellText = Replace(cellText, "**", "")
            cellText = Replace(cellText, "##", "")
            cellText = Trim(cellText)
            
            ' Remove sample text if not replaced
            If InStr(LCase(cellText), "[copilot will") > 0 Or _
               InStr(LCase(cellText), "[detailed test") > 0 Or _
               InStr(LCase(cellText), "[expected") > 0 Or _
               InStr(LCase(cellText), "[integration") > 0 Or _
               InStr(LCase(cellText), "[security") > 0 Or _
               InStr(LCase(cellText), "[error handling") > 0 Or _
               InStr(LCase(cellText), "[invalid input") > 0 Or _
               InStr(LCase(cellText), "[api/system") > 0 Or _
               InStr(LCase(cellText), "[proper error") > 0 Or _
               InStr(LCase(cellText), "[security compliance") > 0 Or _
               InStr(LCase(cellText), "[integration results") > 0 Then
                cellText = ""
            End If
            
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
                        .Font.Bold = True
                        Call ColorTestCategory(tbl.cell(i, j))
                        .ParagraphFormat.Alignment = wdAlignParagraphCenter
                    Case 3 ' Scenario
                        .Font.Color = RGB(52, 53, 65)
                        .ParagraphFormat.Alignment = wdAlignParagraphLeft
                        .Font.Bold = False
                    Case 4 ' Steps
                        .Font.Color = RGB(52, 53, 65)
                        .ParagraphFormat.Alignment = wdAlignParagraphLeft
                        .Font.Size = 8
                    Case 5 ' Expected
                        .Font.Color = RGB(52, 53, 65)
                        .ParagraphFormat.Alignment = wdAlignParagraphLeft
                    Case 6 ' Priority
                        .Font.Bold = True
                        .ParagraphFormat.Alignment = wdAlignParagraphCenter
                        Call ColorPriority(tbl.cell(i, j))
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

' Color test categories
Sub ColorTestCategory(cell As cell)
    Dim text As String
    text = UCase(cell.Range.text)
    
    With cell.Range
        Select Case True
            Case InStr(text, "FUNCTIONAL") > 0
                .Font.Color = RGB(255, 255, 255)
                .Shading.BackgroundPatternColor = RGB(34, 197, 94)
            Case InStr(text, "INTEGRATION") > 0
                .Font.Color = RGB(255, 255, 255)
                .Shading.BackgroundPatternColor = RGB(59, 130, 246)
            Case InStr(text, "SECURITY") > 0
                .Font.Color = RGB(255, 255, 255)
                .Shading.BackgroundPatternColor = RGB(239, 68, 68)
            Case InStr(text, "PERFORMANCE") > 0
                .Font.Color = RGB(255, 255, 255)
                .Shading.BackgroundPatternColor = RGB(245, 158, 11)
            Case InStr(text, "UI") > 0 Or InStr(text, "USER") > 0
                .Font.Color = RGB(255, 255, 255)
                .Shading.BackgroundPatternColor = RGB(139, 92, 246)
            Case InStr(text, "NEGATIVE") > 0
                .Font.Color = RGB(255, 255, 255)
                .Shading.BackgroundPatternColor = RGB(236, 72, 153)
            Case InStr(text, "BOUNDARY") > 0
                .Font.Color = RGB(0, 0, 0)
                .Shading.BackgroundPatternColor = RGB(251, 191, 36)
            Case InStr(text, "ACCESSIBILITY") > 0
                .Font.Color = RGB(255, 255, 255)
                .Shading.BackgroundPatternColor = RGB(16, 185, 129)
            Case Else
                .Font.Color = RGB(0, 0, 0)
                .Shading.BackgroundPatternColor = RGB(229, 231, 235)
        End Select
    End With
End Sub

' Color priority cells (enhanced)
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

' Clear results area while protecting first page
Sub ClearResults()
    Dim doc As Document
    Dim rng As Range
    
    Set doc = ActiveDocument
    
    ' Find the position after the first page content
    Dim clearStart As Long
    clearStart = FirstPageEnd
    
    ' If we have content after the first page, clear it
    If doc.Range.End > clearStart Then
        Set rng = doc.Range(clearStart, doc.Range.End)
        rng.Delete
    End If
End Sub

' Clear entire document
Sub ClearDocument()
    ActiveDocument.Range.Delete
    Set ConfigTable = Nothing
    Set StatusCell = Nothing
    FirstPageEnd = 0
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

' Get JSON value (enhanced parser)
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

' Get long JSON values like descriptions
Function GetJSONValueLong(json As String, key As String) As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim searchStr As String
    Dim braceCount As Integer
    Dim inQuotes As Boolean
    
    GetJSONValueLong = "No description available"
    
    searchStr = """" & key & """:"
    startPos = InStr(json, searchStr)
    
    If startPos > 0 Then
        startPos = startPos + Len(searchStr)
        
        ' Skip whitespace
        Do While Mid(json, startPos, 1) = " "
            startPos = startPos + 1
        Loop
        
        ' Check if value starts with quote
        If Mid(json, startPos, 1) = """" Then
            startPos = startPos + 1
            endPos = startPos
            
            ' Find matching quote, handling escapes
            Do While endPos <= Len(json)
                Dim char As String
                char = Mid(json, endPos, 1)
                
                If char = """" And Mid(json, endPos - 1, 1) <> "\" Then
                    Exit Do
                End If
                endPos = endPos + 1
            Loop
            
            If endPos > startPos Then
                GetJSONValueLong = Mid(json, startPos, endPos - startPos)
            End If
        End If
    End If
End Function

' Format date helper
Function FormatDate(dateStr As String) As String
    On Error Resume Next
    Dim dateVal As Date
    dateVal = CDate(Replace(Left(dateStr, 10), "-", "/"))
    FormatDate = Format(dateVal, "mmm dd, yyyy")
    If Err.Number <> 0 Then FormatDate = Left(dateStr, 10)
    On Error GoTo 0
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
        MsgBox ChrW(9733) & " GitLab connection works!" & vbInformation
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
    
    ' Build CSV with enhanced headers
    For i = 1 To tbl.Rows.count
        For j = 1 To tbl.Columns.count
            On Error Resume Next
            Dim cellText As String
            cellText = tbl.cell(i, j).Range.text
            cellText = Replace(cellText, Chr(13) & Chr(7), "")
            cellText = Replace(cellText, """", """""")
            cellText = Replace(cellText, vbCrLf, " | ")
            csvContent = csvContent & """" & cellText & """"
            If j < tbl.Columns.count Then csvContent = csvContent & ","
            On Error GoTo 0
        Next j
        csvContent = csvContent & vbCrLf
    Next i
    
    ' Save file
    fileName = "ComprehensiveTestCases_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"
    Dim filePath As String
    filePath = Environ("USERPROFILE") & "\Desktop\" & fileName
    
    Open filePath For Output As #1
    Print #1, csvContent
    Close #1
    
    Call SetStatus("Exported comprehensive test cases", "success")
    MsgBox ChrW(9733) & " Comprehensive test cases exported to:" & vbCrLf & _
           "Desktop\" & fileName & vbCrLf & vbCrLf & _
           "Ready for import into test management tools!", vbInformation
End Sub

Sub QuickSetup()
    If ConfigTable Is Nothing Then
        MsgBox "Run StartGitLabGenerator first", vbExclamation
        Exit Sub
    End If
    
    ConfigTable.cell(1, 2).Range.text = "https://gitlab.com"
    ConfigTable.cell(2, 2).Range.text = "Project"
    ConfigTable.cell(3, 2).Range.text = "Epic"
    
    Call SetStatus("Pre-configured for comprehensive GitLab.com analysis", "ready")
End Sub
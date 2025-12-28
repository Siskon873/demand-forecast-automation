Option Explicit

'===========================================================
' DEMAND FORECAST DISTRIBUTION AUTOMATION
' GitHub: https://github.com/your-username/demand-forecast-automation
' License: MIT
' Author: Your Name
' Description: Automates zone-based report generation and distribution
'===========================================================

'===========================================================
' CONFIGURATION - Update these for your environment
'===========================================================
Public Const SEND_EMAILS As Boolean = False              ' Set True to send emails
Public Const PREVIEW_EMAILS As Boolean = True            ' Set True to preview drafts
Public Const USE_PARTIAL_CATEGORY_MATCH As Boolean = True ' Flexible matching

'===========================================================
' MAIN ENTRY POINT
' Press Alt+F8, select RunAutomation, and click Run
'===========================================================
Public Sub RunAutomation()
    On Error GoTo ErrHandler
    Dim startTime As Double: startTime = Timer
    
    ' Initialize
    Call BeginFastMode
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Starting automation..."
    
    ' Get folder paths
    Dim baseFolder As String, sourcesPath As String
    baseFolder = ThisWorkbook.Path
    
    If Len(baseFolder) = 0 Then
        MsgBox "⚠️ Please save this workbook first!", vbExclamation
        GoTo Cleanup
    End If
    
    sourcesPath = baseFolder & "\Sources"
    If Dir(sourcesPath, vbDirectory) = "" Then MkDir sourcesPath
    
    ' Find mapping file
    Application.StatusBar = "Looking for mapping file..."
    Dim mappingPath As String
    mappingPath = FindMappingFile(baseFolder)
    
    If Len(mappingPath) = 0 Then
        MsgBox "❌ Mapping file not found!" & vbCrLf & vbCrLf & _
               "Expected: ZoneWiseMapping.xlsx or similar" & vbCrLf & _
               "Location: " & baseFolder, vbCritical
        GoTo Cleanup
    End If
    
    ' Create logs
    Dim logPath As String
    logPath = sourcesPath & "\Automation_Log_" & Format(Now, "yyyymmdd_hhnnss") & ".txt"
    
    Call WriteLog(logPath, "========================================")
    Call WriteLog(logPath, "AUTOMATION STARTED: " & Now)
    Call WriteLog(logPath, "Demand File: " & ThisWorkbook.Name)
    Call WriteLog(logPath, "Mapping File: " & Dir(mappingPath))
    Call WriteLog(logPath, "========================================")
    
    ' Load mapping
    Application.StatusBar = "Reading zone configuration..."
    Dim wbMap As Workbook
    Set wbMap = Workbooks.Open(mappingPath, ReadOnly:=True, UpdateLinks:=False)
    
    Dim wsMap As Worksheet
    Set wsMap = GetMappingSheet(wbMap)
    
    Dim zoneDict As Object
    Set zoneDict = BuildZoneDict(wsMap)
    wbMap.Close False
    
    Call WriteLog(logPath, "Total Zones: " & zoneDict.Count)
    Call WriteLog(logPath, "")
    
    ' Process zones
    Dim wbDemand As Workbook: Set wbDemand = ThisWorkbook
    Dim filesCreated As Long, emailsSent As Long, totalSheets As Long
    filesCreated = 0: emailsSent = 0: totalSheets = 0
    
    Dim zoneCount As Long: zoneCount = zoneDict.Count
    Dim currentZone As Long: currentZone = 0
    
    Dim zone As Variant
    For Each zone In zoneDict.Keys
        currentZone = currentZone + 1
        Application.StatusBar = "Processing: " & zone & " (" & currentZone & "/" & zoneCount & ")"
        
        Call WriteLog(logPath, "--- Zone: " & zone & " ---")
        
        Dim filePath As String, sheetsInFile As Long
        filePath = ProcessZone(wbDemand, CStr(zone), zoneDict(zone), sourcesPath, logPath, sheetsInFile)
        
        If Len(filePath) > 0 Then
            filesCreated = filesCreated + 1
            totalSheets = totalSheets + sheetsInFile
            Call WriteLog(logPath, "  ✓ File created: " & Dir(filePath))
            
            Dim emailAddr As String
            emailAddr = CStr(zoneDict(zone)("email"))
            If Len(emailAddr) > 0 Then
                Call SendEmail(emailAddr, CStr(zone), filePath, sheetsInFile)
                emailsSent = emailsSent + 1
                Call WriteLog(logPath, "  ✉️ Email prepared: " & emailAddr)
            End If
        Else
            Call WriteLog(logPath, "  ✗ No matching cities")
        End If
        Call WriteLog(logPath, "")
    Next zone
    
    Dim elapsed As Double: elapsed = Round(Timer - startTime, 2)
    
    ' Create summary
    Dim summaryPath As String
    summaryPath = sourcesPath & "\Summary_Report.txt"
    Call CreateSummary(summaryPath, filesCreated, emailsSent, totalSheets, elapsed, zoneCount)
    
    Call WriteLog(logPath, "========================================")
    Call WriteLog(logPath, "COMPLETED: " & Now)
    Call WriteLog(logPath, "Time: " & elapsed & " seconds")
    Call WriteLog(logPath, "========================================")
    
    ' Show results
    Call ShowResults(filesCreated, emailsSent, totalSheets, elapsed, sourcesPath)
    
Cleanup:
    Application.Calculation = xlCalculationAutomatic
    Call EndFastMode
    Application.StatusBar = False
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Call EndFastMode
    Application.StatusBar = False
    MsgBox "❌ Error: " & Err.Description, vbCritical
End Sub

'===========================================================
' FIND MAPPING FILE
'===========================================================
Private Function FindMappingFile(folderPath As String) As String
    FindMappingFile = ""
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folder As Object: Set folder = fso.GetFolder(folderPath)
    Dim file As Object
    
    For Each file In folder.Files
        If InStr(1, file.Name, "Mapping", vbTextCompare) > 0 And _
           (LCase(fso.GetExtensionName(file.Name)) = "xlsx" Or _
            LCase(fso.GetExtensionName(file.Name)) = "xlsm") Then
            FindMappingFile = file.Path
            Exit Function
        End If
    Next file
End Function

'===========================================================
' PROCESS ONE ZONE
'===========================================================
Private Function ProcessZone(wbDemand As Workbook, zoneName As String, zoneInfo As Object, _
                            sourcesPath As String, logPath As String, _
                            ByRef sheetsInFile As Long) As String
    ProcessZone = ""
    sheetsInFile = 0
    
    Dim wbOut As Workbook: Set wbOut = Workbooks.Add
    Application.DisplayAlerts = False
    Do While wbOut.Sheets.Count > 1
        wbOut.Sheets(wbOut.Sheets.Count).Delete
    Loop
    Application.DisplayAlerts = True
    
    Dim cityData As Object: Set cityData = zoneInfo("cities")
    Dim firstSheet As Boolean: firstSheet = True
    
    Dim cityKey As Variant
    For Each cityKey In cityData.Keys
        Dim cityName As String: cityName = CStr(cityKey)
        Dim categoryFilter As String: categoryFilter = CStr(cityData(cityKey))
        
        Call WriteLog(logPath, "  City: " & cityName & " | Filter: " & categoryFilter)
        
        Dim matchedSheet As String
        matchedSheet = FindCitySheet(wbDemand, cityName)
        
        If Len(matchedSheet) > 0 Then
            Dim wsSource As Worksheet: Set wsSource = wbDemand.Sheets(matchedSheet)
            
            If firstSheet Then
                wsSource.Copy Before:=wbOut.Sheets(1)
                Application.DisplayAlerts = False
                wbOut.Sheets(2).Delete
                Application.DisplayAlerts = True
                firstSheet = False
            Else
                wsSource.Copy After:=wbOut.Sheets(wbOut.Sheets.Count)
            End If
            
            Dim wsNew As Worksheet: Set wsNew = wbOut.Sheets(wbOut.Sheets.Count)
            On Error Resume Next
            wsNew.Name = Left(cityName, 31)
            On Error GoTo 0
            
            Dim rowsKept As Long
            rowsKept = FilterByCategory(wsNew, categoryFilter)
            
            If rowsKept > 0 Then
                sheetsInFile = sheetsInFile + 1
                Call WriteLog(logPath, "    ✓ Matched: " & matchedSheet & " (" & rowsKept & " rows)")
            Else
                Application.DisplayAlerts = False
                wsNew.Delete
                Application.DisplayAlerts = True
            End If
        Else
            Call WriteLog(logPath, "    ✗ Sheet not found")
        End If
    Next cityKey
    
    If sheetsInFile > 0 Then
        Dim fileName As String
        fileName = sourcesPath & "\" & zoneName & " Report - " & Format(Date, "dd-MMM-yyyy") & ".xlsx"
        
        On Error Resume Next
        If Dir(fileName) <> "" Then Kill fileName
        On Error GoTo 0
        
        wbOut.SaveAs fileName, FileFormat:=51
        wbOut.Close False
        ProcessZone = fileName
    Else
        wbOut.Close False
    End If
End Function

'===========================================================
' FIND CITY SHEET
'===========================================================
Private Function FindCitySheet(wb As Workbook, cityName As String) As String
    FindCitySheet = ""
    Dim ws As Worksheet
    Dim searchName As String: searchName = LCase(Trim(cityName))
    
    On Error Resume Next
    Set ws = wb.Sheets(cityName)
    If Not ws Is Nothing Then FindCitySheet = ws.Name: Exit Function
    On Error GoTo 0
    
    For Each ws In wb.Worksheets
        Dim sheetName As String: sheetName = LCase(Trim(ws.Name))
        If sheetName = searchName Or InStr(sheetName, searchName) > 0 Then
            FindCitySheet = ws.Name
            Exit Function
        End If
    Next ws
End Function

'===========================================================
' FILTER BY CATEGORY
'===========================================================
Private Function FilterByCategory(ws As Worksheet, categoryFilter As String) As Long
    FilterByCategory = 0
    
    Dim headerRow As Long, categoryCol As Long
    headerRow = FindCategoryColumn(ws, categoryCol)
    
    If headerRow = 0 Or categoryCol = 0 Then
        FilterByCategory = ws.UsedRange.Rows.Count - 1
        Exit Function
    End If
    
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow <= headerRow Then Exit Function
    
    Dim tokens() As String: tokens = Split(categoryFilter, ",")
    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        tokens(i) = Trim(tokens(i))
    Next i
    
    Application.Calculation = xlCalculationManual
    Dim r As Long
    For r = lastRow To headerRow + 1 Step -1
        Dim cellValue As String: cellValue = Trim(CStr(ws.Cells(r, categoryCol).Value))
        If Not MatchesCategory(cellValue, tokens) Then ws.Rows(r).Delete
    Next r
    Application.Calculation = xlCalculationAutomatic
    
    FilterByCategory = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - headerRow
End Function

Private Function MatchesCategory(cellValue As String, tokens() As String) As Boolean
    MatchesCategory = False
    If Len(Trim(cellValue)) = 0 Then Exit Function
    
    Dim i As Long
    For i = LBound(tokens) To UBound(tokens)
        If Len(Trim(tokens(i))) = 0 Then GoTo NextToken
        
        If USE_PARTIAL_CATEGORY_MATCH Then
            If InStr(1, cellValue, tokens(i), vbTextCompare) > 0 Then
                MatchesCategory = True
                Exit Function
            End If
        Else
            If StrComp(cellValue, tokens(i), vbTextCompare) = 0 Then
                MatchesCategory = True
                Exit Function
            End If
        End If
NextToken:
    Next i
End Function

Private Function FindCategoryColumn(ws As Worksheet, ByRef categoryCol As Long) As Long
    FindCategoryColumn = 0
    categoryCol = 0
    
    Dim r As Long, c As Long
    For r = 1 To WorksheetFunction.Min(10, ws.Rows.Count)
        For c = 1 To WorksheetFunction.Min(100, ws.Columns.Count)
            Dim v As String: v = LCase(Trim(CStr(ws.Cells(r, c).Value)))
            If v = "category" Or InStr(v, "categor") > 0 Then
                categoryCol = c
                FindCategoryColumn = r
                Exit Function
            End If
        Next c
    Next r
End Function

'===========================================================
' BUILD ZONE DICTIONARY
'===========================================================
Private Function BuildZoneDict(ws As Worksheet) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim r As Long
    For r = 2 To lastRow
        Dim zoneName As String, cityName As String, category As String, email As String
        zoneName = Trim(CStr(ws.Cells(r, 1).Value))
        cityName = Trim(CStr(ws.Cells(r, 2).Value))
        category = Trim(CStr(ws.Cells(r, 3).Value))
        email = Trim(CStr(ws.Cells(r, 4).Value))
        
        If Len(zoneName) = 0 Or Len(cityName) = 0 Then GoTo NextRow
        
        If Not dict.Exists(zoneName) Then
            Dim zInfo As Object: Set zInfo = CreateObject("Scripting.Dictionary")
            Set zInfo("cities") = CreateObject("Scripting.Dictionary")
            zInfo("email") = email
            dict.Add zoneName, zInfo
        End If
        
        If Not dict(zoneName)("cities").Exists(cityName) Then
            dict(zoneName)("cities").Add cityName, category
        End If
NextRow:
    Next r
    
    Set BuildZoneDict = dict
End Function

Private Function GetMappingSheet(wb As Workbook) As Worksheet
    On Error Resume Next
    Set GetMappingSheet = wb.Worksheets("Normalized")
    If GetMappingSheet Is Nothing Then Set GetMappingSheet = wb.Worksheets(1)
    On Error GoTo 0
End Function

'===========================================================
' EMAIL
'===========================================================
Private Sub SendEmail(emailList As String, zoneName As String, filePath As String, sheetsCount As Long)
    On Error Resume Next
    
    Dim OutApp As Object, OutMail As Object
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    Dim msg As String
    msg = "Dear Team," & vbCrLf & vbCrLf & _
          "Please find attached the report for " & zoneName & "." & vbCrLf & vbCrLf & _
          "Report contains " & sheetsCount & " city sheet(s)" & vbCrLf & _
          "Date: " & Format(Date, "dd-MMM-yyyy") & vbCrLf & vbCrLf & _
          "Regards," & vbCrLf & "Automated System"
    
    With OutMail
        .To = emailList
        .Subject = zoneName & " - Report - " & Format(Date, "dd-MMM-yyyy")
        .Body = msg
        If Dir(filePath) <> "" Then .Attachments.Add filePath
        
        If SEND_EMAILS Then
            .Send
        ElseIf PREVIEW_EMAILS Then
            .Display
        End If
    End With
    
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

'===========================================================
' REPORTING
'===========================================================
Private Sub CreateSummary(filePath As String, filesCreated As Long, emailsSent As Long, _
                         totalSheets As Long, elapsed As Double, totalZones As Long)
    Dim fNum As Integer: fNum = FreeFile
    Open filePath For Output As #fNum
    
    Print #fNum, "AUTOMATION SUMMARY"
    Print #fNum, "=================="
    Print #fNum, ""
    Print #fNum, "Date: " & Format(Date, "dd-MMM-yyyy")
    Print #fNum, "Time: " & Format(Now, "hh:mm:ss AM/PM")
    Print #fNum, ""
    Print #fNum, "Results:"
    Print #fNum, "  Processing Time: " & elapsed & " seconds"
    Print #fNum, "  Zones Processed: " & totalZones
    Print #fNum, "  Files Created: " & filesCreated
    Print #fNum, "  Total Sheets: " & totalSheets
    Print #fNum, "  Emails Sent: " & emailsSent
    Print #fNum, ""
    Print #fNum, "Success Rate: " & Format((filesCreated / totalZones) * 100, "0.0") & "%"
    
    Close #fNum
End Sub

Private Sub ShowResults(filesCreated As Long, emailsSent As Long, totalSheets As Long, _
                       elapsed As Double, sourcesPath As String)
    Dim msg As String
    msg = "✅ AUTOMATION COMPLETE!" & vbCrLf & vbCrLf
    msg = msg & "Files: " & filesCreated & " | Sheets: " & totalSheets & vbCrLf
    msg = msg & "Emails: " & emailsSent & " | Time: " & elapsed & "s" & vbCrLf & vbCrLf
    msg = msg & "Output: " & sourcesPath & vbCrLf & vbCrLf
    msg = msg & "Open folder?"
    
    If MsgBox(msg, vbInformation + vbYesNo) = vbYes Then
        Shell "explorer.exe """ & sourcesPath & """", vbNormalFocus
    End If
End Sub

Private Sub WriteLog(logPath As String, msg As String)
    On Error Resume Next
    Dim fNum As Integer: fNum = FreeFile
    Open logPath For Append As #fNum
    Print #fNum, msg
    Close #fNum
End Sub

'===========================================================
' PERFORMANCE HELPERS
'===========================================================
Private Sub BeginFastMode()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
End Sub

Private Sub EndFastMode()
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
End Sub
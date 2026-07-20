Option Explicit
' Destination tabs in PB_Summary_NEW.
Private Const DEST_LIE As String = "rga_YYYYMMDD_lie TOD"
Private Const DEST_ABV As String = "RGA0300_YYYYMMDD_ABV TOD"
Public Sub Import_RGA_TOD_Reports()
    Dim hostWb As Workbook
    Dim lieWb As Workbook
    Dim abvWb As Workbook
    Dim lieWs As Worksheet
    Dim abvWs As Worksheet
    Dim lieDest As Worksheet
    Dim abvDest As Worksheet
    Dim lieData As Range
    Dim abvData As Range
    Dim downloadFolder As String
    Dim lieFile As String
    Dim abvFile As String
    Dim expectedDate As String
    Dim reportDate As String
    Dim abvDate As String
    Dim oldCalculation As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldAutomationSecurity As Long
    Dim settingsChanged As Boolean
    Dim completed As Boolean
    Dim errorText As String
    Dim lieRows As Long
    Dim lieCols As Long
    Dim abvRows As Long
    Dim abvCols As Long
    On Error GoTo ImportFailed
    Set hostWb = ThisWorkbook
    downloadFolder = Environ$("USERPROFILE") & "\Downloads\"
    If Not FolderExists(downloadFolder) Then
        downloadFolder = hostWb.Path & Application.PathSeparator
    End If
    ' Find the newest LIE report in Downloads.
    lieFile = FindNewestFile(downloadFolder, "rga_????????_lie.*")
    ' Use the previous Monday-Friday business day, not the newest download.
    expectedDate = Format$(PreviousBusinessDay(Date), "yyyymmdd")
    lieFile = FindNewestFile(downloadFolder, _
                             "rga_" & expectedDate & "_lie.*")
    If Len(lieFile) = 0 Then
        lieFile = PickReport(downloadFolder, _
                             "Select the rga_YYYYMMDD_lie report")
        lieFile = PickReport(downloadFolder, "Select LIE report for " & expectedDate)
    End If
    If Len(lieFile) = 0 Then Exit Sub
    reportDate = ExtractReportDate(lieFile)
    If Len(reportDate) = 0 Then
        Err.Raise vbObjectError + 1001, "Import_RGA_TOD_Reports", _
                  "The LIE filename does not contain an 8-digit YYYYMMDD date: " & _
                  FileNameOnly(lieFile)
    End If
    ' Use the ABV report from the same report date.
    abvFile = FindNewestFile(downloadFolder, _
                             "rga0300_" & reportDate & "_abv.*")
    If Len(abvFile) = 0 Then
        abvFile = PickReport(downloadFolder, _
                             "Select the rga0300_" & reportDate & "_abv report")
    End If
    If Len(abvFile) = 0 Then Exit Sub
    abvDate = ExtractReportDate(abvFile)
    If abvDate <> reportDate Then
        Err.Raise vbObjectError + 1002, "Import_RGA_TOD_Reports", _
                  "The two reports have different dates." & vbCrLf & _
                  "LIE: " & reportDate & vbCrLf & _
                  "ABV: " & abvDate
    End If
    Set lieDest = GetWorksheet(hostWb, DEST_LIE)
    Set abvDest = GetWorksheet(hostWb, DEST_ABV)
    If lieDest Is Nothing Then
        Err.Raise vbObjectError + 1003, "Import_RGA_TOD_Reports", _
                  "Destination tab not found: " & DEST_LIE
    End If
    If abvDest Is Nothing Then
        Err.Raise vbObjectError + 1004, "Import_RGA_TOD_Reports", _
                  "Destination tab not found: " & DEST_ABV
    End If
    oldCalculation = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldAutomationSecurity = Application.AutomationSecurity
    settingsChanged = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.AutomationSecurity = 3 ' msoAutomationSecurityForceDisable
    Application.StatusBar = "Opening RGA reports..."
    Set lieWb = Workbooks.Open(Filename:=lieFile, _
                               UpdateLinks:=0, _
                               ReadOnly:=True, _
                               AddToMru:=False, _
                               IgnoreReadOnlyRecommended:=True, _
                               Notify:=False, _
                               Local:=True)
    Set abvWb = Workbooks.Open(Filename:=abvFile, _
                               UpdateLinks:=0, _
                               ReadOnly:=True, _
                               AddToMru:=False, _
                               IgnoreReadOnlyRecommended:=True, _
                               Notify:=False, _
                               Local:=True)
    Set lieWs = lieWb.Worksheets(1)
    Set abvWs = abvWb.Worksheets(1)
    Set lieData = RealDataRange(lieWs)
    Set abvData = RealDataRange(abvWs)
    If lieData Is Nothing Then
        Err.Raise vbObjectError + 1005, "Import_RGA_TOD_Reports", _
                  "The LIE report is empty: " & FileNameOnly(lieFile)
    End If
    If abvData Is Nothing Then
        Err.Raise vbObjectError + 1006, "Import_RGA_TOD_Reports", _
                  "The ABV report is empty: " & FileNameOnly(abvFile)
    End If
    lieRows = lieData.Rows.Count
    lieCols = lieData.Columns.Count
    abvRows = abvData.Rows.Count
    abvCols = abvData.Columns.Count
    Application.StatusBar = "Replacing TOD tab data..."
    ReplaceSheetData lieData, lieDest
    ReplaceSheetData abvData, abvDest
    lieDest.Calculate
    abvDest.Calculate
    completed = True
CleanUp:
    On Error Resume Next
    Application.CutCopyMode = False
    If Not lieWb Is Nothing Then lieWb.Close SaveChanges:=False
    If Not abvWb Is Nothing Then abvWb.Close SaveChanges:=False
    If settingsChanged Then
        Application.AutomationSecurity = oldAutomationSecurity
        Application.Calculation = oldCalculation
        Application.DisplayAlerts = oldDisplayAlerts
        Application.EnableEvents = oldEnableEvents
        Application.ScreenUpdating = oldScreenUpdating
    End If
    Application.StatusBar = False
    On Error GoTo 0
    If completed Then
        MsgBox "Import completed for " & reportDate & "." & vbCrLf & vbCrLf & _
               FileNameOnly(lieFile) & vbCrLf & _
               "  -> " & DEST_LIE & " (" & lieRows & " rows x " & lieCols & " columns)" & vbCrLf & vbCrLf & _
               FileNameOnly(abvFile) & vbCrLf & _
               "  -> " & DEST_ABV & " (" & abvRows & " rows x " & abvCols & " columns)", _
               vbInformation, "PB Summary import"
    ElseIf Len(errorText) > 0 Then
        MsgBox errorText, vbExclamation, "PB Summary import"
    End If
    Exit Sub
ImportFailed:
    errorText = "Import stopped: " & Err.Description
    Resume CleanUp
End Sub
Private Function PreviousBusinessDay(ByVal referenceDate As Date) As Date
    Dim resultDate As Date
    resultDate = DateValue(referenceDate) - 1
    Do While Weekday(resultDate, vbMonday) > 5
        resultDate = resultDate - 1
    Loop
    PreviousBusinessDay = resultDate
End Function
Private Sub ReplaceSheetData(ByVal sourceData As Range, _
                             ByVal destinationSheet As Worksheet)
    If destinationSheet.ProtectContents Then
        Err.Raise vbObjectError + 1010, "ReplaceSheetData", _
                  "Destination tab is protected: " & destinationSheet.Name
    End If
    ' Clear old values/formulas while retaining the template's layout and objects.
    destinationSheet.UsedRange.ClearContents
    ' Paste report results, not formulas or links back to the downloaded file.
    sourceData.Copy
    destinationSheet.Range("A1").PasteSpecial _
        Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False
End Sub
Private Function RealDataRange(ByVal ws As Worksheet) As Range
    Dim lastRowCell As Range
    Dim lastColumnCell As Range
    Set lastRowCell = ws.Cells.Find(What:="*", _
                                    After:=ws.Cells(1, 1), _
                                    LookIn:=xlFormulas, _
                                    LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, _
                                    SearchDirection:=xlPrevious, _
                                    MatchCase:=False, _
                                    SearchFormat:=False)
    If lastRowCell Is Nothing Then Exit Function
    Set lastColumnCell = ws.Cells.Find(What:="*", _
                                       After:=ws.Cells(1, 1), _
                                       LookIn:=xlFormulas, _
                                       LookAt:=xlPart, _
                                       SearchOrder:=xlByColumns, _
                                       SearchDirection:=xlPrevious, _
                                       MatchCase:=False, _
                                       SearchFormat:=False)
    Set RealDataRange = ws.Range(ws.Cells(1, 1), _
                                 ws.Cells(lastRowCell.Row, lastColumnCell.Column))
End Function
Private Function GetWorksheet(ByVal wb As Workbook, _
                              ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function
Private Function FindNewestFile(ByVal folderPath As String, _
                                ByVal fileMask As String) As String
    Dim fileName As String
    Dim fullPath As String
    Dim newestTime As Date
    Dim currentTime As Date
    On Error GoTo SearchFinished
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    fileName = Dir$(folderPath & fileMask, _
                    vbNormal Or vbHidden Or vbReadOnly Or vbSystem Or vbArchive)
    Do While Len(fileName) > 0
        fullPath = folderPath & fileName
        currentTime = FileDateTime(fullPath)
        If Len(FindNewestFile) = 0 Or currentTime > newestTime Then
            FindNewestFile = fullPath
            newestTime = currentTime
        End If
        fileName = Dir$()
    Loop
SearchFinished:
End Function
Private Function PickReport(ByVal initialFolder As String, _
                            ByVal dialogTitle As String) As String
    With Application.FileDialog(3) ' msoFileDialogFilePicker
        .Title = dialogTitle
        .AllowMultiSelect = False
        .InitialFileName = initialFolder
        .Filters.Clear
        .Filters.Add "Excel and CSV reports", _
                     "*.xlsx;*.xls;*.xlsm;*.xlsb;*.csv"
        .Filters.Add "All files", "*.*"
        If .Show = -1 Then PickReport = .SelectedItems(1)
    End With
End Function
Private Function ExtractReportDate(ByVal fullPath As String) As String
    Dim fileName As String
    Dim candidate As String
    Dim i As Long
    Dim yearNumber As Long
    Dim monthNumber As Long
    Dim dayNumber As Long
    fileName = FileNameOnly(fullPath)
    For i = 1 To Len(fileName) - 7
        candidate = Mid$(fileName, i, 8)
        If candidate Like "########" Then
            yearNumber = CLng(Left$(candidate, 4))
            monthNumber = CLng(Mid$(candidate, 5, 2))
            dayNumber = CLng(Right$(candidate, 2))
            If yearNumber >= 2000 And yearNumber <= 2099 _
               And monthNumber >= 1 And monthNumber <= 12 _
               And dayNumber >= 1 And dayNumber <= 31 Then
                ExtractReportDate = candidate
                Exit Function
            End If
        End If
    Next i
End Function
Private Function FileNameOnly(ByVal fullPath As String) As String
    Dim slashPosition As Long
    Dim forwardSlashPosition As Long
    slashPosition = InStrRev(fullPath, "\")
    forwardSlashPosition = InStrRev(fullPath, "/")
    If forwardSlashPosition > slashPosition Then slashPosition = forwardSlashPosition
    FileNameOnly = Mid$(fullPath, slashPosition + 1)
End Function
Private Function FolderExists(ByVal folderPath As String) As Boolean
    FolderExists = CreateObject("Scripting.FileSystemObject").FolderExists(folderPath)
End Function

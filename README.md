Option Explicit

' Source reports dated on the previous business day.
Private Const DEST_RGA_LIE As String = "rga_YYYYMMDD_lie TOD"
Private Const DEST_RGA_ABV As String = "RGA0300_YYYYMMDD_ABV TOD"
Private Const DEST_RBC_INTAM As String = "rbc_YYYYMMDD_intam"
Private Const DEST_RBC_SLFSM As String = "rbc_YYYYMMDD_slfsm TOD"
Private Const DEST_PHN_INTAM As String = "phn_YYYYMMDD_intam"
Private Const DEST_PHN_SLFSM As String = "phn_YYYYMMDD_slfsm TOD"

' Source reports dated today.
Private Const DEST_SHORT_STOCK As String = "YYYYMMDD_SHORT_STOCK_DETAIL"
Private Const DEST_INTEREST As String = "YYYYMMDD_INTEREST_ACTIVITY"

Public Sub Import_All_PB_Reports()
    Const REPORT_COUNT As Long = 8

    Dim hostWb As Workbook
    Dim targets(1 To REPORT_COUNT) As Worksheet
    Dim labels(1 To REPORT_COUNT) As String
    Dim masks(1 To REPORT_COUNT) As String
    Dim targetNames(1 To REPORT_COUNT) As String
    Dim reportFiles(1 To REPORT_COUNT) As String
    Dim results(1 To REPORT_COUNT) As String

    Dim downloadFolder As String
    Dim expectedPreviousDate As String
    Dim previousReportDate As String
    Dim todayReportDate As String
    Dim i As Long
    Dim summaryText As String

    Dim oldCalculation As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldAutomationSecurity As Long
    Dim settingsChanged As Boolean
    Dim completed As Boolean
    Dim errorText As String

    On Error GoTo ImportFailed

    Set hostWb = ThisWorkbook
    downloadFolder = Environ$("USERPROFILE") & "\Downloads\"
    If Not FolderExists(downloadFolder) Then
        downloadFolder = hostWb.Path & Application.PathSeparator
    End If

    expectedPreviousDate = Format$(PreviousBusinessDay(Date), "yyyymmdd")
    todayReportDate = Format$(Date, "yyyymmdd")

    labels(1) = "RGA LIE"
    masks(1) = "rga_" & expectedPreviousDate & "_lie.*"
    targetNames(1) = DEST_RGA_LIE

    ' Resolve LIE first. Its date becomes the common date for RGA/RBC/PHN.
    reportFiles(1) = ResolveReport(downloadFolder, masks(1), labels(1), expectedPreviousDate)
    If Len(reportFiles(1)) = 0 Then Exit Sub

    previousReportDate = ExtractReportDate(reportFiles(1))
    If Len(previousReportDate) = 0 Then
        Err.Raise vbObjectError + 2001, "Import_All_PB_Reports", _
                  "No YYYYMMDD date was found in: " & FileNameOnly(reportFiles(1))
    End If

    labels(2) = "RGA ABV"
    masks(2) = "rga0300_" & previousReportDate & "_abv.*"
    targetNames(2) = DEST_RGA_ABV

    labels(3) = "RBC INTAM"
    masks(3) = "rbc90050_" & previousReportDate & "_intam.*"
    targetNames(3) = DEST_RBC_INTAM

    labels(4) = "RBC SLFSM TOD"
    masks(4) = "rbc90050_" & previousReportDate & "_slfsm.*"
    targetNames(4) = DEST_RBC_SLFSM

    labels(5) = "PHN INTAM"
    masks(5) = "phn01226_" & previousReportDate & "_intam.*"
    targetNames(5) = DEST_PHN_INTAM

    labels(6) = "PHN SLFSM TOD"
    masks(6) = "phn01226_" & previousReportDate & "_slfsm.*"
    targetNames(6) = DEST_PHN_SLFSM

    labels(7) = "SHORT STOCK DETAIL"
    masks(7) = todayReportDate & "_SHORT_STOCK_DETAIL.*"
    targetNames(7) = DEST_SHORT_STOCK

    labels(8) = "INTEREST ACTIVITY"
    masks(8) = todayReportDate & "_INTEREST_ACTIVITY.*"
    targetNames(8) = DEST_INTEREST

    ' Locate every source before changing any destination tab.
    For i = 2 To REPORT_COUNT
        If i <= 6 Then
            reportFiles(i) = ResolveReport(downloadFolder, masks(i), labels(i), previousReportDate)
        Else
            reportFiles(i) = ResolveReport(downloadFolder, masks(i), labels(i), todayReportDate)
        End If

        If Len(reportFiles(i)) = 0 Then Exit Sub
    Next i

    ' Confirm all destination tabs before importing.
    For i = 1 To REPORT_COUNT
        Set targets(i) = GetWorksheetFlexible(hostWb, targetNames(i))
        If targets(i) Is Nothing Then
            Err.Raise vbObjectError + 2002, "Import_All_PB_Reports", _
                      "Destination tab not found: " & targetNames(i)
        End If
    Next i

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
    Application.AutomationSecurity = 3

    For i = 1 To REPORT_COUNT
        Application.StatusBar = "Importing " & labels(i) & " (" & i & "/" & REPORT_COUNT & ")..."
        results(i) = ImportOneReport(reportFiles(i), targets(i))
    Next i

    completed = True

CleanUp:
    On Error Resume Next
    Application.CutCopyMode = False

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
        summaryText = "Import completed." & vbCrLf & _
                      "Previous-business-day reports: " & previousReportDate & vbCrLf & _
                      "Same-day reports: " & todayReportDate & vbCrLf & vbCrLf

        For i = 1 To REPORT_COUNT
            summaryText = summaryText & labels(i) & " -> " & _
                          targets(i).Name & " (" & results(i) & ")" & vbCrLf
        Next i

        MsgBox summaryText, vbInformation, "PB Summary import"
    ElseIf Len(errorText) > 0 Then
        MsgBox errorText, vbExclamation, "PB Summary import"
    End If

    Exit Sub

ImportFailed:
    errorText = "Import stopped: " & Err.Description
    Resume CleanUp
End Sub

Private Function ImportOneReport(ByVal sourcePath As String, _
                                 ByVal destinationSheet As Worksheet) As String
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim sourceData As Range
    Dim rowCount As Long
    Dim columnCount As Long
    Dim savedErrorNumber As Long
    Dim savedErrorText As String

    On Error GoTo ImportOneFailed

    If destinationSheet.ProtectContents Then
        Err.Raise vbObjectError + 2010, "ImportOneReport", _
                  "Destination tab is protected: " & destinationSheet.Name
    End If

    Set sourceWb = Workbooks.Open(Filename:=sourcePath, _
                                  UpdateLinks:=0, _
                                  ReadOnly:=True, _
                                  AddToMru:=False, _
                                  IgnoreReadOnlyRecommended:=True, _
                                  Notify:=False, _
                                  Local:=True)

    Set sourceWs = sourceWb.Worksheets(1)
    Set sourceData = RealDataRange(sourceWs)

    If sourceData Is Nothing Then
        Err.Raise vbObjectError + 2011, "ImportOneReport", _
                  "Source report is empty: " & FileNameOnly(sourcePath)
    End If

    rowCount = sourceData.Rows.Count
    columnCount = sourceData.Columns.Count

    destinationSheet.UsedRange.ClearContents
    sourceData.Copy
    destinationSheet.Range("A1").PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False

    ImportOneReport = rowCount & " rows x " & columnCount & " columns"
    sourceWb.Close SaveChanges:=False
    Exit Function

ImportOneFailed:
    savedErrorNumber = Err.Number
    savedErrorText = Err.Description
    On Error Resume Next
    Application.CutCopyMode = False
    If Not sourceWb Is Nothing Then sourceWb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise savedErrorNumber, "ImportOneReport", savedErrorText
End Function

Private Function ResolveReport(ByVal folderPath As String, _
                               ByVal fileMask As String, _
                               ByVal reportLabel As String, _
                               ByVal reportDate As String) As String
    ResolveReport = FindNewestFile(folderPath, fileMask)

    If Len(ResolveReport) = 0 Then
        ResolveReport = PickReport(folderPath, _
                                   "Select " & reportLabel & " report for " & reportDate)
    End If
End Function

Private Function PreviousBusinessDay(ByVal referenceDate As Date) As Date
    Dim resultDate As Date

    resultDate = DateValue(referenceDate) - 1
    Do While Weekday(resultDate, vbMonday) > 5
        resultDate = resultDate - 1
    Loop

    PreviousBusinessDay = resultDate
End Function

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

Private Function GetWorksheetFlexible(ByVal wb As Workbook, _
                                      ByVal expectedName As String) As Worksheet
    Dim ws As Worksheet
    Dim normalizedExpected As String

    On Error Resume Next
    Set GetWorksheetFlexible = wb.Worksheets(expectedName)
    On Error GoTo 0

    If Not GetWorksheetFlexible Is Nothing Then Exit Function

    normalizedExpected = NormalizeName(expectedName)
    For Each ws In wb.Worksheets
        If NormalizeName(ws.Name) = normalizedExpected Then
            Set GetWorksheetFlexible = ws
            Exit Function
        End If
    Next ws
End Function

Private Function NormalizeName(ByVal textValue As String) As String
    Dim result As String

    result = LCase$(Trim$(textValue))
    result = Replace(result, "_", "")
    result = Replace(result, " ", "")
    result = Replace(result, "-", "")
    NormalizeName = result
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
    With Application.FileDialog(3)
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

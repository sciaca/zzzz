Option Explicit

' Source reports dated on the previous business day.
Private Const DEST_RGA_LIE As String = "rga_YYYYMMDD_lie TOD"
Private Const DEST_RGA_ABV As String = "RGA0300_YYYYMMDD_ABV TOD"
Private Const DEST_RBC_INTAM As String = "rbc_YYYYMMDD_intam"
Private Const DEST_RBC_SLFSM As String = "rbc_YYYYMMDD_slfsm TOD"
Private Const DEST_PHN_INTAM As String = "phn_YYYYMMDD_intam"
Private Const DEST_PHN_SLFSM As String = "phn_YYYYMMDD_slfsm TOD"

' Source report dated today and appended as history.
Private Const DEST_SHORT_STOCK As String = "YYYYMMDD_SHORT_STOCK_DETAIL"

Public Sub Import_All_PB_Reports()
    Const REPORT_COUNT As Long = 7

    Dim hostWb As Workbook
    Dim targets(1 To REPORT_COUNT) As Worksheet
    Dim labels(1 To REPORT_COUNT) As String
    Dim masks(1 To REPORT_COUNT) As String
    Dim targetNames(1 To REPORT_COUNT) As String
    Dim reportFiles(1 To REPORT_COUNT) As String
    Dim results(1 To REPORT_COUNT) As String
    Dim appendMode(1 To REPORT_COUNT) As Boolean
    Dim skipSourceHeader(1 To REPORT_COUNT) As Boolean

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

    expectedPreviousDate = Format$(PreviousCanadianBusinessDay(Date), "yyyymmdd")
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
    appendMode(7) = True
    skipSourceHeader(7) = True

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
        results(i) = ImportOneReport(reportFiles(i), targets(i), _
                                     appendMode(i), skipSourceHeader(i))
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
                                 ByVal destinationSheet As Worksheet, _
                                 ByVal appendData As Boolean, _
                                 ByVal skipHeader As Boolean) As String
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim sourceData As Range
    Dim dataToPaste As Range
    Dim pasteCell As Range
    Dim rowCount As Long
    Dim columnCount As Long
    Dim destinationLastRow As Long
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

    If skipHeader Then
        ' Clean any repeated source header left by an earlier import.
        RemoveRepeatedHeaderRows destinationSheet, sourceData

        If sourceData.Rows.Count <= 1 Then
            ImportOneReport = "source contains header only; 0 rows appended"
            sourceWb.Close SaveChanges:=False
            Exit Function
        End If

        Set dataToPaste = sourceData.Offset(1, 0).Resize( _
                              sourceData.Rows.Count - 1, _
                              sourceData.Columns.Count)
    Else
        Set dataToPaste = sourceData
    End If

    If appendData Then
        destinationLastRow = LastUsedRow(destinationSheet)

        If destinationLastRow > 0 Then
            If BlockAlreadyAtBottom(dataToPaste, destinationSheet, destinationLastRow) Then
                ImportOneReport = "same data already present; 0 rows appended"
                sourceWb.Close SaveChanges:=False
                Exit Function
            End If

            Set pasteCell = destinationSheet.Cells(destinationLastRow + 1, 1)
        Else
            Set pasteCell = destinationSheet.Range("A1")
        End If
    Else
        destinationSheet.UsedRange.ClearContents
        Set pasteCell = destinationSheet.Range("A1")
    End If

    rowCount = dataToPaste.Rows.Count
    columnCount = dataToPaste.Columns.Count

    dataToPaste.Copy
    pasteCell.PasteSpecial Paste:=xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False

    If appendData Then
        ImportOneReport = rowCount & " rows appended"
        If skipHeader Then ImportOneReport = ImportOneReport & "; source header removed"
    Else
        ImportOneReport = rowCount & " rows x " & columnCount & " columns replaced"
    End If

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

Private Sub RemoveRepeatedHeaderRows(ByVal destinationSheet As Worksheet, _
                                     ByVal sourceData As Range)
    Dim destinationLastRow As Long
    Dim rowNumber As Long
    Dim columnNumber As Long
    Dim rowMatchesHeader As Boolean

    destinationLastRow = LastUsedRow(destinationSheet)

    For rowNumber = destinationLastRow To 2 Step -1
        rowMatchesHeader = True

        For columnNumber = 1 To sourceData.Columns.Count
            If HeaderValue(destinationSheet.Cells(rowNumber, columnNumber).Value2) <> _
               HeaderValue(sourceData.Cells(1, columnNumber).Value2) Then
                rowMatchesHeader = False
                Exit For
            End If
        Next columnNumber

        If rowMatchesHeader Then destinationSheet.Rows(rowNumber).Delete
    Next rowNumber
End Sub

Private Function HeaderValue(ByVal cellValue As Variant) As String
    If IsError(cellValue) Then
        HeaderValue = "#ERROR"
    ElseIf IsEmpty(cellValue) Then
        HeaderValue = vbNullString
    Else
        HeaderValue = LCase$(Trim$(CStr(cellValue)))
    End If
End Function

Private Function BlockAlreadyAtBottom(ByVal sourceData As Range, _
                                      ByVal destinationSheet As Worksheet, _
                                      ByVal destinationLastRow As Long) As Boolean
    Dim existingBlock As Range
    Dim sourceValues As Variant
    Dim existingValues As Variant
    Dim rowNumber As Long
    Dim columnNumber As Long
    Dim firstExistingRow As Long

    If destinationLastRow < sourceData.Rows.Count Then Exit Function

    firstExistingRow = destinationLastRow - sourceData.Rows.Count + 1
    Set existingBlock = destinationSheet.Cells(firstExistingRow, 1).Resize( _
                            sourceData.Rows.Count, sourceData.Columns.Count)

    sourceValues = sourceData.Value2
    existingValues = existingBlock.Value2

    If sourceData.Cells.CountLarge = 1 Then
        BlockAlreadyAtBottom = ValuesMatch(sourceValues, existingValues)
        Exit Function
    End If

    For rowNumber = 1 To sourceData.Rows.Count
        For columnNumber = 1 To sourceData.Columns.Count
            If Not ValuesMatch(sourceValues(rowNumber, columnNumber), _
                               existingValues(rowNumber, columnNumber)) Then
                Exit Function
            End If
        Next columnNumber
    Next rowNumber

    BlockAlreadyAtBottom = True
End Function

Private Function ValuesMatch(ByVal firstValue As Variant, _
                             ByVal secondValue As Variant) As Boolean
    If IsError(firstValue) Or IsError(secondValue) Then
        ValuesMatch = (IsError(firstValue) And IsError(secondValue))
    ElseIf IsEmpty(firstValue) And IsEmpty(secondValue) Then
        ValuesMatch = True
    Else
        ValuesMatch = (CStr(firstValue) = CStr(secondValue))
    End If
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range

    Set lastCell = ws.Cells.Find(What:="*", _
                                 After:=ws.Cells(1, 1), _
                                 LookIn:=xlFormulas, _
                                 LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, _
                                 SearchDirection:=xlPrevious, _
                                 MatchCase:=False, _
                                 SearchFormat:=False)

    If Not lastCell Is Nothing Then LastUsedRow = lastCell.Row
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

Private Function PreviousCanadianBusinessDay(ByVal referenceDate As Date) As Date
    Dim resultDate As Date

    resultDate = DateValue(referenceDate) - 1
    Do While Weekday(resultDate, vbMonday) > 5 _
             Or IsCanadianBankHoliday(resultDate)
        resultDate = resultDate - 1
    Loop

    PreviousCanadianBusinessDay = resultDate
End Function

Private Function IsCanadianBankHoliday(ByVal dateToCheck As Date) As Boolean
    Dim calendarDate As Date
    Dim calendarYear As Long
    Dim christmasObserved As Date
    Dim boxingDayObserved As Date

    calendarDate = DateValue(dateToCheck)
    calendarYear = Year(calendarDate)

    ' Bank of Canada / Ontario banking calendar.
    If calendarDate = ObservedWeekday(DateSerial(calendarYear, 1, 1)) Then GoTo HolidayFound
    If calendarDate = NthWeekdayOfMonth(calendarYear, 2, vbMonday, 3) Then GoTo HolidayFound
    If calendarDate = EasterSunday(calendarYear) - 2 Then GoTo HolidayFound
    If calendarDate = VictoriaDay(calendarYear) Then GoTo HolidayFound
    If calendarDate = ObservedWeekday(DateSerial(calendarYear, 7, 1)) Then GoTo HolidayFound
    If calendarDate = NthWeekdayOfMonth(calendarYear, 8, vbMonday, 1) Then GoTo HolidayFound
    If calendarDate = NthWeekdayOfMonth(calendarYear, 9, vbMonday, 1) Then GoTo HolidayFound
    If calendarDate = ObservedWeekday(DateSerial(calendarYear, 9, 30)) Then GoTo HolidayFound
    If calendarDate = NthWeekdayOfMonth(calendarYear, 10, vbMonday, 2) Then GoTo HolidayFound
    If calendarDate = ObservedWeekday(DateSerial(calendarYear, 11, 11)) Then GoTo HolidayFound

    christmasObserved = ObservedWeekday(DateSerial(calendarYear, 12, 25))
    boxingDayObserved = ObservedWeekday(DateSerial(calendarYear, 12, 26))

    Do While boxingDayObserved = christmasObserved _
             Or Weekday(boxingDayObserved, vbMonday) > 5
        boxingDayObserved = boxingDayObserved + 1
    Loop

    If calendarDate = christmasObserved Then GoTo HolidayFound
    If calendarDate = boxingDayObserved Then GoTo HolidayFound

    Exit Function

HolidayFound:
    IsCanadianBankHoliday = True
End Function

Private Function ObservedWeekday(ByVal holidayDate As Date) As Date
    Select Case Weekday(holidayDate, vbMonday)
        Case 6
            ObservedWeekday = holidayDate + 2
        Case 7
            ObservedWeekday = holidayDate + 1
        Case Else
            ObservedWeekday = holidayDate
    End Select
End Function

Private Function NthWeekdayOfMonth(ByVal calendarYear As Long, _
                                   ByVal calendarMonth As Long, _
                                   ByVal weekdayNumber As VbDayOfWeek, _
                                   ByVal occurrenceNumber As Long) As Date
    Dim firstDate As Date
    Dim dayOffset As Long

    firstDate = DateSerial(calendarYear, calendarMonth, 1)
    dayOffset = (weekdayNumber - Weekday(firstDate, vbSunday) + 7) Mod 7

    NthWeekdayOfMonth = firstDate + dayOffset + (7 * (occurrenceNumber - 1))
End Function

Private Function VictoriaDay(ByVal calendarYear As Long) As Date
    Dim resultDate As Date

    resultDate = DateSerial(calendarYear, 5, 24)
    Do While Weekday(resultDate, vbMonday) <> 1
        resultDate = resultDate - 1
    Loop

    VictoriaDay = resultDate
End Function

Private Function EasterSunday(ByVal calendarYear As Long) As Date
    Dim a As Long
    Dim b As Long
    Dim c As Long
    Dim d As Long
    Dim e As Long
    Dim f As Long
    Dim g As Long
    Dim h As Long
    Dim i As Long
    Dim k As Long
    Dim l As Long
    Dim m As Long
    Dim easterMonth As Long
    Dim easterDay As Long

    a = calendarYear Mod 19
    b = calendarYear \ 100
    c = calendarYear Mod 100
    d = b \ 4
    e = b Mod 4
    f = (b + 8) \ 25
    g = (b - f + 1) \ 3
    h = (19 * a + b - d - g + 15) Mod 30
    i = c \ 4
    k = c Mod 4
    l = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = (a + 11 * h + 22 * l) \ 451
    easterMonth = (h + l - 7 * m + 114) \ 31
    easterDay = ((h + l - 7 * m + 114) Mod 31) + 1

    EasterSunday = DateSerial(calendarYear, easterMonth, easterDay)
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

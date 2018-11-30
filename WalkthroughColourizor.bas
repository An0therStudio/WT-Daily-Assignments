Attribute VB_Name = "WalkthroughColourizor"
'Base output path
Const PATH_OUT = "\\YGK01CFP01\Operations\Call Execution\WT Team\Assignments\"
'Const PATH_OUT = "\\YGK01CFP01\Operations\Call Execution\WT Team\Test\"
'Template file to copy
Const FILE_SRC = "AllCalls Week of.xlsx"
'Path of Template file
Const PATH_SRC = "\\YGK01CFP01\Operations\Call Execution\WT Team\WT Checklists & Templates\"
Sub WalkthroughColourizor()
    Const CANCELLED = "Cancelled"
    Const COMPLETED = "Completed"
    Const THIRD = "3rd"
    Const MC = "MC HEBERT"
    Const HEBERT = "MARIE CLAUDE HEBERT"
    Const LANGDON = "HEATHER LANGDON"
    Dim r As Range
    
    'If CheckIfSheetExists("Temp") Then _
        Application.DisplayAlerts = False _
        Sheets("Temp").Delete _
        Application.DisplayAlerts = True _
    End If
'    ActiveWorkbook.SaveAs Filename:=("\\YGK01CFP01\Operations\Call Execution\WT Team\Assignments\Daily Files\" & ActiveWorkbook.Name & Format(CStr(Now), "yyy_mm_dd_hh_mm") & ".xlsx")
    
    Application.ScreenUpdating = False
    
    Call ExtractData

    Sheets("Temp").Select
    Range("A1").Select
    Set mylastcell = Cells(1, 1).SpecialCells(xlLastCell)
    mylastcelladd = Cells(mylastcell.Row, mylastcell.column).Address
    myrange = "A1:" & mylastcelladd
    Set r = Range(myrange) '.Select
    
    r.Value = Application.Trim(r)
    
    'Delete Sub-Total line (last row)
    'Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).EntireRow.Delete
    
    'Delete row if search term is found in designated column
    Call DeleteRows("H", CANCELLED)
    Call DeleteRows("F", COMPLETED)
    Call DeleteRows("F", THIRD)
    Call DeleteRows("C", MC)
    Call DeleteRows("C", HEBERT)
    Call DeleteRows("C", LANGDON)
    Call DeleteRows("D", MC)
    Call DeleteRows("D", HEBERT)
    Call DeleteRows("D", LANGDON)
    
    Call CreateAssignmentsSheet
    
    Application.ScreenUpdating = True

    
End Sub

Private Sub ExtractData()
    Dim wkbIn, wkbOut, sheetIn, sheetOut As String
    wkbIn = ActiveWorkbook.name
    wkbOut = ActiveWorkbook.name
    sheetIn = "Report 1"
    sheetOut = "Temp"
    'Create destination worksheet
    With ActiveWorkbook
        .Sheets.Add(After:=.Sheets(.Sheets.Count)).name = "Temp"
    End With
    
    'Copy Date
    'Call CopyColumnTo("C5", wkbIn, sheetIn, "A1", wkbOut, sheetOut)
    Call CopyColumnTo("E5", wkbIn, sheetIn, "A1", wkbOut, sheetOut)
    'Format Date
    Set r = ActiveSheet.Range("A2")
    ActiveSheet.Range(r, ActiveSheet.Cells(Rows.Count, r.column).End(xlUp).Address).NumberFormat = "mm-dd hh:mm"

    'Copy Company Name
    'Call CopyColumnTo("G5", wkbIn, sheetIn, "B1", wkbOut, sheetOut)
    Call CopyColumnTo("C5", wkbIn, sheetIn, "B1", wkbOut, sheetOut)
    
    'Copy Leader Name
    'Call CopyColumnTo("I5", wkbIn, sheetIn, "C1", wkbOut, sheetOut)
    Call CopyColumnTo("J5", wkbIn, sheetIn, "C1", wkbOut, sheetOut)
    
    'Copy Assistant Name
    'Call CopyColumnTo("J5", wkbIn, sheetIn, "D1", wkbOut, sheetOut)
    Call CopyColumnTo("G5", wkbIn, sheetIn, "D1", wkbOut, sheetOut)
    
    'Copy Conference ID
    'Call CopyColumnTo("K5", wkbIn, sheetIn, "E1", wkbOut, sheetOut)
    Call CopyColumnTo("D5", wkbIn, sheetIn, "E1", wkbOut, sheetOut)
    
    'Copy WT Status
    'Call CopyColumnTo("P5", wkbIn, sheetIn, "F1", wkbOut, sheetOut)
    Call CopyColumnTo("M5", wkbIn, sheetIn, "F1", wkbOut, sheetOut)
    
    'Copy Ace Bridge
    'Call CopyColumnTo("Q5", wkbIn, sheetIn, "G1", wkbOut, sheetOut)
    Call CopyColumnTo("H5", wkbIn, sheetIn, "G1", wkbOut, sheetOut)
    
    'Copy Reservation Status
    'Call CopyColumnTo("R5", wkbIn, sheetIn, "H1", wkbOut, sheetOut)
    Call CopyColumnTo("P5", wkbIn, sheetIn, "H1", wkbOut, sheetOut)
    
    'Copy Company Number
    'Call CopyColumnTo("F5", wkbIn, sheetIn, "I1", wkbOut, sheetOut)
    Call CopyColumnTo("B5", wkbIn, sheetIn, "I1", wkbOut, sheetOut)
    
    'Copy Owner Number
    'Call CopyColumnTo("F5", wkbIn, sheetIn, "I1", wkbOut, sheetOut)
    Call CopyColumnTo("Q5", wkbIn, sheetIn, "J1", wkbOut, sheetOut)
    
    
    
End Sub

Private Sub CopyColumn(startCell As String)
    Dim r As Range
    Set r = Sheet1.Range(startCell)
    'Copy all used cells in column below startCell(inclusive)
    Sheet1.Range(r, Sheet1.Cells(Rows.Count, r.column).End(xlUp).Address).Copy
End Sub

Private Sub PasteColumn(startCell As String)
    ActiveSheet.paste Destination:=Worksheets("Temp").Range(startCell)
End Sub

Private Sub CopyColumnTo(copyStart, copyBook, copySheet, pasteStart, pasteBook, pasteSheet)
    Dim r As Range
    Dim lookupWB As Workbook
    
    Set lookupWB = Workbooks(pasteBook)
    'Copy and paste all used cells in column below copyStart(inclusive)
    Set r = Workbooks(copyBook).Worksheets(copySheet).Range(copyStart)
    Workbooks(copyBook).Worksheets(copySheet).Range(r, Workbooks(copyBook).Worksheets(copySheet).Cells(Rows.Count, r.column).End(xlUp).Address).Copy _
        Destination:=lookupWB.Worksheets(pasteSheet).Range(pasteStart)
    
End Sub

Private Sub DeleteRows(column As String, target As String)
    
    Dim myTarget As String
    myTarget = target
    
    Dim Rng As Range
    Dim i As Long, j As Long
    
    ' Calc last row number
    j = Range(column & Rows.Count).End(xlUp).Row
    
    ' Collect rows with MyTarget
    For i = j To 1 Step -1
        'Set RngTrim = Rows(i)
        If WorksheetFunction.CountIf(Rows(i), myTarget) > 0 Then
            If Rng Is Nothing Then
                Set Rng = Rows(i)
            Else
                Set Rng = Union(Rng, Rows(i))
            End If
        End If
    Next
    
    ' Delete rows with MyTarget
    If Not Rng Is Nothing Then Rng.Delete
    
    ' Update UsedRange
    With ActiveSheet.UsedRange: End With
  
End Sub

Private Sub CreateAssignmentsSheet()
    Dim wkbOut, wkbSrc As Workbook
    Dim dteToday As Date
    Dim pathOut, fileOut, strDayOfTheWeek, strMonday, strYear, strMonthNumber, strMonthName, strMonthFolder As String
    
    'Capture currently used workbook
    Set wkbSrc = Workbooks(ActiveWorkbook.name)
    'String name of workbook
    thisbook = CStr(wkbSrc.name)
    
    'Format today's date and construct strings for output file
    dteToday = Format(Date, "mm/dd/yyyy")
    'Current Year
    strYear = Format(dteToday, "yyyy")
    'Numeric Current Month
    strMonthNumber = Month(dteToday)
    'String Current Month
    strMonthName = MonthName(Month(dteToday))
    'Month Folder Name
    'strMonthFolder = strMonthNumber & "_" & strMonthName
    'Monday's date
    strMonday = dteToday - Weekday(dteToday) + 2
    'Monday's month #
    strMonMonthNumber = Month(strMonday)
    'Monday's month name
    strMonMonthName = MonthName(Month(strMonday))
    'Month Folder Name
    strMonthFolder = strMonMonthNumber & "_" & strMonMonthName
    strMonday = MonthName(Month(strMonday)) & " " & Format(strMonday, "dd")
    'String Today's Weekday
    strDayOfTheWeek = WeekdayName(Weekday(dteToday))
    
    'Output file
    fileOut = "AllCalls Week of " & strMonday & ".xlsx"
    
    'Check for Source file
    CheckDirectory (PATH_SRC)
    
    'Check for Base directory
    CheckDirectory (PATH_OUT)
    'Check for Year directory
    pathOut = PATH_OUT & strYear & "\"
    CheckDirectory (pathOut)
    'Check for Month directory
    pathOut = pathOut & strMonthFolder & "\"
    CheckDirectory (pathOut)
    
    'Create weekly file if it does not exist
    If Not FileExists(pathOut & fileOut) Then
        Call FileCopy(PATH_SRC & FILE_SRC, pathOut & fileOut)
    End If
    'Open weekly file
    Set wkbOut = Workbooks.Open(Filename:=pathOut & fileOut)
    
    'Copy Temp page to correct day of the week
    Call CopyColumnTo("A1", thisbook, "Temp", "A2", fileOut, strDayOfTheWeek)
    Call CopyColumnTo("B1", thisbook, "Temp", "B2", fileOut, strDayOfTheWeek)
    Call CopyColumnTo("E1", thisbook, "Temp", "C2", fileOut, strDayOfTheWeek)
    Call CopyColumnTo("F1", thisbook, "Temp", "D2", fileOut, strDayOfTheWeek)
    Call CopyColumnTo("G1", thisbook, "Temp", "E2", fileOut, strDayOfTheWeek)
    Call CopyColumnTo("J1", thisbook, "Temp", "J2", fileOut, strDayOfTheWeek)
    Call CopyColumnTo("I1", thisbook, "Temp", "K2", fileOut, strDayOfTheWeek)
    Call CopyColumnTo("D1", thisbook, "Temp", "L2", fileOut, strDayOfTheWeek)
    
    Call ColourizeRows(CStr(wkbOut.name), strDayOfTheWeek)
    
    'Save destination file after copying
    wkbOut.Save
    'Delete Temp sheet after copying
    Application.DisplayAlerts = False
    wkbSrc.Sheets("Temp").Delete
    Application.DisplayAlerts = True
        

End Sub

Private Sub CheckDirectory(strDir)
    If Dir(strDir, vbDirectory) = "" Then
            MkDir strDir
    End If
End Sub

Function FileExists(fullFileName As String) As Boolean
    FileExists = VBA.Len(VBA.Dir(fullFileName)) > 0
End Function

Function CheckIfSheetExists(SheetName As String) As Boolean
    Dim ws As Worksheet
      IsExists = False
      For Each ws In Worksheets
        If SheetName = ws.name Then
          IsExists = True
          Exit Function
        End If
      Next ws
End Function

Private Sub ColourizeRows(wkbName, day)
    Dim difference, workDays As Integer
    'Activate target worksheet
    Workbooks(wkbName).Sheets(day).Activate
    ' Select cell A1, *first line of data*.
      Range("A1").Select
      ' Set Do loop to stop when an empty cell is reached.
      Do Until IsEmpty(ActiveCell)
        difference = 0
        workDays = 0
        'Calculate number of days between date macro is run and date of call
        difference = DateDiff("d", Date, CDate(ActiveCell.Value))
        'Adjust difference if call is on weekend to ensure same colouring as Friday
        If Weekday(CDate(ActiveCell.Value)) = 7 Then
            difference = difference - 1
        ElseIf Weekday(CDate(ActiveCell.Value)) = 1 Then
            difference = difference - 2
        End If
        'Colour based on difference value and adjust after weekend.
        workDays = WorksheetFunction.NetworkDays(Date, CDate(ActiveCell.Value))
        If ((difference - workDays) >= 0) Then
            difference = difference - (difference - workDays + 1)
        End If
        Select Case difference
            'Orange
            Case 0
                ActiveCell.Resize(1, 5).Interior.Color = RGB(228, 106, 10)
            'Pink
            Case 1
                ActiveCell.Resize(1, 5).Interior.Color = RGB(220, 150, 150)
            'Green
            Case 2
                ActiveCell.Resize(1, 5).Interior.Color = RGB(120, 150, 60)
            'Blue
            Case 3
                ActiveCell.Resize(1, 5).Interior.Color = RGB(85, 140, 210)
            'Purple
            Case Else
                ActiveCell.Resize(1, 5).Interior.Color = RGB(153, 153, 255)
        End Select
         ' Step down 1 row from present location.
         ActiveCell.Offset(1, 0).Select
      Loop
End Sub

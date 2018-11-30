VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAssignments 
   Caption         =   "Assignments Generator"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9360
   OleObjectBlob   =   "frmAssignments.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAssignments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim FilterContents() As String
Dim TeamContents() As String
Dim arrayFilters() As Variant
Dim lastTab As Integer
Const FILE_PATH As String = "\\YGK01CFP01\Operations\Call Execution\WT Team\Macros & Tools\DailyAssignments\AssignmentsGenerator\Data Files\"
'Const FILE_PATH As String = "\\YGK01CFP01\Operations\Call Execution\WT Team\Test\"
Const FILE_NAME As String = "WTAG.dat"
Const SEPERATOR As String = ";"

Private Sub UserForm_Initialize()
    Dim blArrayPopulated As Boolean
    
    'MultiPageMain.Value = 0
    arrayFilters = Array()
    Call LoadControls
    blArrayPopulated = LoadTextFile
    lastTab = 0
    If blArrayPopulated Then FilterContents = arrayFilters(0)
    MultiPageMain.Value = 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        SaveToTextFile
        ' Tip: If you want to prevent closing UserForm by Close (×) button in the right-top corner of the UserForm, just uncomment the following line:
        ' Cancel = True
    End If
End Sub

'**********************************USER INTERFACE***********************************

Private Sub MultiPageMain_Change()
    Dim intSize As Integer
    'Save list boxes to arrays
    'MsgBox ("Changed. Where are we? Value: " & MultiPageMain.Value)
    Select Case MultiPageMain.Value
        'Opening Main. Save other tabs. Populate Current tab.
        Case 0
            Call SaveFilters
            Call ClearForm
            Call LoadControls
            Call LoadTeam
            Call LoadFilters
            'MsgBox ("Tab: " & tabTeam.Value)
            
        'Opening Filters. Save other tabs. Populate Current tab.
        Case 1
            Call SaveTeam
            Call ClearForm
            Call LoadControls
            'FilterContents = arrayFilters(0)
            Call LoadFilters
            Call CreateTabs
            'Ensure first tab is loaded
            'tabTeam.Value = 0
            
            
        'Opening AS Accounts. Save other tabs. Populate Current tab.
        Case 2
            Call SaveTeam
            Call SaveFilters
            Call ClearForm
            Call LoadControls
            'TODO: CAPTURE ARRAYS
            
        'Error has occured. Page does not exist or new page needs to be defined
        Case Default
            MsgBox ("Error. Page Not Defined.")
    End Select
    'Call ClearForm
    'Call UserForm_Initialize
    
End Sub

Private Sub tabTeam_Change()
    'Save current tab
    Call SaveFilters
    'Clear and load new tab
    Call ClearForm
    Call LoadControls
    Call LoadTeam
    If tabTeam.Value >= 0 Then
        FilterContents = arrayFilters(tabTeam.Value)
    Else
        Erase FilterContents
    End If
    Call LoadFilters
    lastTab = tabTeam.Value
    'MsgBox ("Tab: " & tabTeam.Value)
End Sub

Private Sub lbFilters_Click()
    lbHeaders.Value = ""
    btnAddFilter.Enabled = False
    btnRemoveFilter.Enabled = True
End Sub

Private Sub lbHeaders_Click()
    lbFilters.Value = ""
    btnRemoveFilter.Enabled = False
    btnAddFilter.Enabled = True
End Sub

Private Sub lbTeam_Click()
    btnDelete.Enabled = True
    If Len(tbName.Text) = 0 Then
        btnAdd.Enabled = False
    End If
End Sub

Private Sub btnAdd_Click()
    Dim blnFound As Boolean
    blnFound = True
    If Len(Trim(tbName.Value)) > 0 Then
        Dim intSize As Integer
        intSize = lbTeam.ListCount - 1
        'Ensure list was not empty
        If intSize >= 0 Then
            'Check for duplicates
            For i = 0 To intSize
                If StrComp(lbTeam.List(i), tbName.Text) = 0 Then
                    Call MsgBox("Username already exists!", vbCritical, "DUPLICATE ENTRY")
                    blnFound = True
                    Exit For
                Else
                    blnFound = False
                End If
            Next i
        End If
        If blnFound = False Or intSize < 0 Then
            lbTeam.AddItem UCase((tbName.Text))
            tbName.Value = ""
            btnAdd.Enabled = False
            'New name has been added. Increase counter.
            intSize = intSize + 1
            'Add a placeholder in Array of Arrays to ensure AoA index matches listbox index
            ReDim Preserve arrayFilters(0 To intSize)
            Erase FilterContents
            arrayFilters(intSize) = FilterContents
            'SaveToTextFile
        End If
    End If
End Sub

Private Sub btnDelete_Click()
    Dim strName As String
    Dim blnFound As Boolean
    Dim index As Integer
    
    blnFound = False
    strName = lbTeam.Text
    index = 0
    
    'Delete selected item
    If lbTeam.ListIndex >= 0 Then
        i = MsgBox("Are you sure you wish to delete user '" & strName & "'?", vbYesNo, "CONFIRM REMOVAL")
        'If user presses OK
        If i = 6 Then
            index = lbTeam.ListIndex
            
            'Check to see if a tab has been created
            For j = 0 To tabTeam.Tabs.Count - 1
                If tabTeam.Tabs(j).Caption = strName Then
                    'blnFound = False
                    blnFound = True
                    'Exit loop if tab is found
                    Exit For
                Else
                    'blnFound = True
                    blnFound = False
                    
                End If
            Next j
            'Remove tab if it has been created
            If blnFound Then
                tabTeam.Tabs.Remove (lbTeam.List(index))
                'SaveToTextFile
            End If
            'Update Array of Arrays
            Call DeleteMemberFromAoA(index)
            lbTeam.RemoveItem (index)
            'Disable button if last item was removed
            If lbTeam.ListCount <= 0 Then
                btnDelete.Enabled = False
                Erase TeamContents
                Erase FilterContents
                lbFilters.Clear
            End If
        'Call DeleteMemberFromAoA(index)
        End If
    End If
End Sub

Private Sub btnAddFilter_Click()
Dim criteria As String
    'Check to ensure there is a selection
    If lbHeaders.ListIndex >= 0 Then
        'Limited input for Call Date is accepted
        If StrComp(lbHeaders.Text, "Call Date") = 0 Then
            Dim day As Integer
            'Set day to useless value
            day = 100
            'Loop until valid input is received
            Do
                criteria = InputBox("Select an offset: " & vbNewLine & "0 = Same Day | 1 = One Day | 2 = Two Days | 3 = Three Days | 4 = Four Days" & vbNewLine & "Enter Criteria: ", lbHeaders.Text & " Filter")
                'Check for a number
                If Len(criteria) > 0 Then
                    If IsNumeric(criteria) Then
                        day = criteria
                    End If
                'Break Loop on cancel
                Else
                    Exit Do
                End If
            Loop While day > 4
        'Any input accepted for other filters
        Else
            criteria = InputBox("Enter Criteria: ", lbHeaders.Text & " Filter")
        End If
        'Add filter and criteria to listbox
        If Len(criteria) > 0 Then
            With lbFilters
                .AddItem
                .List(lbFilters.ListCount - 1, 0) = lbHeaders.Text
                .List(lbFilters.ListCount - 1, 1) = criteria
            End With
        End If
    End If
End Sub

Private Sub btnRemoveFilter_Click()
    'Delete selected item
    If lbFilters.ListIndex >= 0 Then
        lbFilters.RemoveItem (lbFilters.ListIndex)
        If lbFilters.ListCount <= 0 Then
            'Disable button if last item was removed
            btnRemoveFilter.Enabled = False
            Erase FilterContents
            arrayFilters(lastTab) = FilterContents
        End If
    End If
End Sub

Private Sub tbName_Change()
    'Toggle buttons when name is entered
    btnAdd.Enabled = True
    btnDelete.Enabled = False
    lbTeam.Value = ""
End Sub

Private Sub ClearForm()
    'Clear controls
    tbName.Value = ""
    lbTeam.Clear
    lbFilters.Clear
    lbHeaders.Clear
End Sub

Private Sub LoadControls()
    btnAddFilter.Enabled = False
    btnRemoveFilter.Enabled = False
    btnAdd.Enabled = False
    btnDelete.Enabled = False
    With lbHeaders
        .AddItem "Call Date"
        .AddItem "Company Name"
        .AddItem "Conference ID"
        .AddItem "WT Status"
        .AddItem "Bridge"
        .IntegralHeight = False
    End With
    With lbFilters
        .IntegralHeight = False
        .BoundColumn = 1
        .TextColumn = 2
    End With
    With lbTeam
        .IntegralHeight = False
    End With
End Sub

Private Sub SaveFilters()
    Dim intSize As Integer
    intSize = lbFilters.ListCount - 1
    'Ensure list was not empty
    If intSize >= 0 Then
        ReDim FilterContents(0 To intSize, 0 To 1) As String
        For i = 0 To intSize
            'Fill multi-dimension array
            For j = 0 To 1
                FilterContents(i, j) = lbFilters.List(i, j)
            Next j
        Next i
        'Save current tab to array of arrays
        'tab order is the same as array order. Use tab position to determine array position
        If (Not TeamContents) = -1 Then
            'Team Array is empty
            teamSize = -1
        Else
            'Get number of rows in array
            teamSize = UBound(TeamContents, 1)
        End If
        If teamSize >= 0 Then
            arrayFilters(lastTab) = FilterContents
        End If
    'Reset FilterContents if list empty
    'Else
        'Erase FilterContents
        'arrayFilters(lastTab) = FilterContents
    End If
End Sub

Private Sub SaveTeam()
    Dim intSize As Integer
    intSize = lbTeam.ListCount - 1
    'Ensure list was not empty
    If intSize >= 0 Then
        ReDim TeamContents(0 To intSize) As String
        'Save list to array
        For i = 0 To intSize
            TeamContents(i) = lbTeam.List(i)
        Next i
    End If
End Sub

Private Sub LoadFilters()
    Dim intSize As Integer
    
    'Ensure filter contents has been loaded
    If (Not TeamContents) = -1 Then
        'Team Array is empty
        teamSize = -1
    Else
        'Get number of rows in array
        teamSize = UBound(TeamContents, 1)
    End If
    If teamSize >= 0 Then
        If tabTeam.Value >= 0 Then
            FilterContents = arrayFilters(tabTeam.Value)
        Else
            FilterContents = arrayFilters(0)
        End If
    End If
    
    'Check for empty array
    If (Not FilterContents) = -1 Then
        'Array is empty
        intSize = -1
    Else
        'Get number of rows in array
        intSize = UBound(FilterContents, 1)
    End If
    'If array is not empty
    If intSize >= 0 Then
        'Add each item to the list
        For i = 0 To intSize
            With lbFilters
                .AddItem
                .List(lbFilters.ListCount - 1, 0) = FilterContents(i, 0)
                .List(lbFilters.ListCount - 1, 1) = FilterContents(i, 1)
            End With
        Next i
    End If
End Sub

Private Sub LoadTeam()
    Dim intSize As Integer
    'Check for empty array
    If (Not TeamContents) = -1 Then
        'Array is empty
        intSize = -1
    Else
        'Get number of rows in array
        intSize = UBound(TeamContents, 1)
    End If
    If intSize >= 0 Then
        'Add each item to the list
        For i = 0 To intSize
            With lbTeam
                .AddItem (TeamContents(i))
            End With
        Next i
    End If
End Sub

Private Sub DeleteMemberFromAoA(index As Integer)
    Dim size As Integer
    size = UBound(arrayFilters)
    'Is index last in list?
    If index = size Then
        'Does list have only 1 item?
        If size > 0 Then
            'Erase last item
            ReDim Preserve arrayFilters(0 To (size - 1))
        Else
            'List is empty, erase array
            Erase arrayFilters
        End If
    Else
        For i = index To size - 1
            arrayFilters(i) = arrayFilters(i + 1)
        Next i
        ReDim Preserve arrayFilters(0 To (size - 1))
    End If
End Sub

Private Sub CreateTabs()
    Dim intSize As Integer
    'Check for empty array
    If (Not TeamContents) = -1 Then
        'Array is empty
        intSize = -1
    Else
        'Get number of rows in array
        intSize = UBound(TeamContents, 1)
    End If
    If intSize >= 0 Then
        For i = 0 To intSize
            Dim blnFound As Boolean
            blnFound = True
            'Look for existing tab
            For j = 0 To tabTeam.Tabs.Count - 1
                If tabTeam.Tabs(j).Caption = TeamContents(i) Then
                    blnFound = True
                    'Exit loop if tab is found
                    Exit For
                Else
                    blnFound = False
                End If
            Next j
            'If not found or tab collection is empty
            If blnFound = False Or (tabTeam.Tabs.Count - 1) < 0 Then
                tabTeam.Tabs.Add (TeamContents(i))
                tabTeam.Tabs(i).Caption = TeamContents(i)
            End If
        Next i
    End If
End Sub

Private Sub SaveToTextFile()
    Dim strLine As String
    Dim fileNum, teamSize, filterSize, intSize, i As Integer
    
    'save list boxes to array before writing to file
    Call SaveTeam
    Call SaveFilters
    i = 0
    'get next available file number
    fileNum = FreeFile
    Open FILE_PATH & FILE_NAME For Output Access Write As #fileNum
    'Check for empty array
    If (Not TeamContents) = -1 Then
        'Team Array is empty
        teamSize = -1
    Else
        'Get number of rows in array
        teamSize = UBound(TeamContents, 1)
    End If
    If teamSize >= 0 Then
        'Loop through each record and write the line to the file
        Do
            strLine = ""
            'strLine = strLine & TeamContents(i) & SEPERATOR
            strLine = strLine & TeamContents(i)
            'remove last seperator
            'strLine = Left(strLine, Len(strLine) - Len(SEPERATOR))
            Print #fileNum, strLine
            'Loop through array of arrays
            'For j = 0 To UBound(arrayFilters)
                'TODO
                'FilterContents = arrayFilters(j)
                FilterContents = arrayFilters(i)
                'Check for empty array
                If (Not FilterContents) = -1 Then
                'Filter Array is empty
                    filterSize = -1
                Else
                    'Get number of rows in array
                    filterSize = UBound(FilterContents, 1)
                End If
                'clear line
                strLine = ""
                'If team member has filters
                If filterSize >= 0 Then
                    For t = 0 To filterSize
                        strLine = strLine & FilterContents(t, 0) & SEPERATOR & FilterContents(t, 1) & SEPERATOR
                    Next t
                    'remove last seperator
                    strLine = Left(strLine, Len(strLine) - Len(SEPERATOR))
                End If
                Print #fileNum, strLine
            'Next j
            i = i + 1
        Loop While i <= teamSize
    End If
    Close #fileNum
End Sub


Private Function LoadTextFile() As Boolean
    Dim strLine, arrLine() As String
    Dim fileNum, teamSize, filterSize, j As Integer
    Dim blHeader, blTeamFound As Boolean
    
    'array 0-based indexing
    teamSize = -1
    filterSize = -1
    'Flag to determine if line being read is a header line or detail line
    blHeader = True
    'Flag to determine if array has been populated
    blTeamFound = False
    'Get next available file number
    fileNum = FreeFile
    'Check if file exists
    If (VBA.Len(VBA.Dir(FILE_PATH & FILE_NAME)) > 0) Then
        Open FILE_PATH & FILE_NAME For Input Access Read As #fileNum
        While Not EOF(fileNum)
            'Read one line at a time
            Line Input #fileNum, strLine
            'Alternate between header (name) and data (filter) lines
            If blHeader Then
                lbTeam.AddItem (strLine)
                teamSize = teamSize + 1
                blHeader = False
            Else
                j = 0
                'Split csv line into array
                arrLine = Split(strLine, SEPERATOR)
                If UBound(arrLine) >= 0 Then
                    'Resize (and empty) FilterContents. Add one to ensure result of UBOUND is divisible by 2. _
                        Minus one for 0-based index
                    ReDim FilterContents(0 To (((UBound(arrLine) + 1) / 2) - 1), 0 To 1)
                    For i = 0 To UBound(arrLine)
                        'Array contents are key-value pairs
                        'Key
                        FilterContents(j, 0) = arrLine(i)
                        'Value
                        FilterContents(j, 1) = arrLine(i + 1)
                        'Increment i by 2 on every loop
                        filterSize = filterSize + 1
                        i = i + 1
                        j = j + 1
                    Next i
                Else
                    Erase FilterContents
                End If
                If teamSize >= 0 Then
                    blTeamFound = True
                    ReDim Preserve arrayFilters(0 To teamSize) As Variant
                    arrayFilters(teamSize) = FilterContents
                Else
                    blTeamFound = False
                    Call MsgBox("Error loading file. Application will exit.", vbCritical, "Critical Error")
                    End
                End If
                blHeader = True
            End If
        Wend
    End If
    Close #fileNum
    Call SaveTeam
    LoadTextFile = blTeamFound
End Function

'****************************************ASSIGNMENTS LOGIC***********************************************

Private Sub btnCreateAssignments_Click()
    Dim wkbIn, strDayOfTheWeek, strFilter, strAgent As String
    Dim dteToday As Date
    Dim teamSize, filterSize, intLowIndex, intLowCount, intAssigned, intTotal, intNew As Integer
    
    dteToday = Format(Date, "mm/dd/yyyy")
    
    wkbIn = ActiveWorkbook.name
    'Check for empty array
    If (Not TeamContents) = -1 Then
        'Array is empty
        teamSize = -1
    Else
        'Get number of rows in array
        teamSize = UBound(arrayFilters)
        ReDim intCount(0 To teamSize) As Integer
    End If
    
    If teamSize >= 0 Then
        'Step through each team member
        For i = 0 To teamSize
            intCount(i) = 0
            FilterContents = arrayFilters(i)
            If (Not FilterContents) = -1 Then
                'Array is empty
                filterSize = -1
            Else
                'Get number of rows in array
                filterSize = UBound(FilterContents)
            End If
            'Step through each filter
            For j = 0 To filterSize
                strFilter = FilterContents(j, 1)
                strAgent = TeamContents(i)
                'TODO: Skip weekends
                'Determine if filter needs to be converted to a date
                    Select Case strFilter
                    'Orange
                    Case 0
                        strFilter = dteToday
                    'Pink
                    Case 1
                        strFilter = dteToday + 1
                    'Green
                    Case 2
                        strFilter = dteToday + 2
                    'Blue
                    Case 3
                        strFilter = dteToday + 3
                    'Purple
                    Case 4
                        strFilter = dteToday + 4
                    End Select
                'Assign according to agent and filter info
                intCount(i) = intCount(i) + ApplyFilter(strFilter, strAgent)
            Next j
        Next i
        'Find team member with least number of WT
        intTotal = GetLastLine() - 1
        intAssigned = 0
        'tally previously assigned calls
        For i = 0 To teamSize
            intAssigned = intAssigned + intCount(i)
        Next i
        Do While intTotal > intAssigned
            intLowCount = intCount(0)
            intLowIndex = 0
            For i = 0 To teamSize
                If intCount(i) < intLowCount Then
                    intLowIndex = i
                    intLowCount = intCount(i)
                End If
            Next i
            strAgent = TeamContents(intLowIndex)
            intNew = GetNextLine(strAgent)
            intCount(intLowIndex) = intCount(intLowIndex) + intNew
            intAssigned = intAssigned + intNew
        Loop
    Else
        MsgBox ("Warning: Please add at least one team member before generating!")
    End If
End Sub

Private Function ApplyFilter(ByVal filter As String, ByVal userName As String) As Integer
    Dim myTarget, myName, strCompany As String
    Dim Rng As Range
    Dim i As Long, j As Long
    Dim blnDate As Boolean
    Dim dteLine As Date
    Dim intCount, intAssigned As Integer
    myTarget = filter
    myName = userName
    Set Rng = Nothing
    intCount = 0
    'dteLine = ""

    ' Calc last row number
    j = GetLastLine()
    
    'If search criterea is a date
    If IsDate(myTarget) Then
        For i = j To 1 Step -1
            If Rng Is Nothing Then
                Set Rng = Rows(i)
                'Check for existing assignment
                If Rng.Columns(6) = "" Then
                    'Avoid type mismatch
                    If IsDate(Rng.Columns(1)) Then
                        'Alter date formatting
                        dteLine = Rng.Columns(1)
                        arrDate = Split(dteLine, " ")
                        dteLine = arrDate(0)
                        'Assign if the line matches
                        If CDate(myTarget) = dteLine Then
                            Rng.Columns(6) = myName
                            intCount = intCount + 1
                            strCompany = Rng.Columns(2)
                            'Call next procedure (with coercion)
                            intCount = intCount + AssignBuddies((myName), strCompany)
                            strCompany = ""
                        End If
                    End If
                End If
                Set Rng = Nothing
            End If
        Next
    Else
        ' Find rows with MyTarget
        For i = j To 1 Step -1
        'If criteria is found apply name and reset range
            If WorksheetFunction.CountIf(Rows(i), myTarget) > 0 Then
                If Rng Is Nothing Then
                    Set Rng = Rows(i)
                    If Rng.Columns(6) = "" Then
                        Rng.Columns(6) = myName
                        intCount = intCount + 1
                        strCompany = Rng.Columns(2)
                        'Call next procedure (with coercion)
                        intCount = intCount + AssignBuddies((myName), strCompany)
                        strCompany = ""
                    End If
                    Set Rng = Nothing
                End If
            End If
        Next
    End If
    ' Update UsedRange
    With ActiveSheet.UsedRange: End With
    ApplyFilter = intCount
End Function

Private Function AssignBuddies(userName As String, companyName As String) As Integer
    Dim myTarget, myName As String
    Dim Rng As Range
    Dim i As Long, j As Long
    Dim intCount As Integer
    myTarget = companyName
    myName = userName
    Set Rng = Nothing
    intCount = 0
    ' Calc last row number
    j = GetLastLine()
    ' Find rows with MyTarget
    For i = j To 1 Step -1
    'If criteria is found apply name and reset range
        If WorksheetFunction.CountIf(Rows(i), myTarget) > 0 Then
            If Rng Is Nothing Then
                Set Rng = Rows(i)
                If Rng.Columns(6) = "" Then
                    Rng.Columns(6) = myName
                    intCount = intCount + 1
                End If
                Set Rng = Nothing
            End If
        End If
    Next
    ' Update UsedRange
    With ActiveSheet.UsedRange: End With
    AssignBuddies = intCount
End Function

Private Function GetNextLine(name As String) As Integer
    Dim i As Long, j As Long
    Dim Rng As Range
    Dim intCount As Integer
    Dim strAgent, strCompany As String
    strAgent = name
    intCount = 0
    j = GetLastLine()
    Set Rng = Nothing
    For i = j To 1 Step -1
        If Rng Is Nothing Then
            Set Rng = Rows(i)
            If Rng.Columns(6) = "" Then
                intCount = intCount + 1
                Rng.Columns(6) = strAgent
                strCompany = Rng.Columns(2)
                intCount = intCount + AssignBuddies((strAgent), strCompany)
                Exit For
            End If
            Set Rng = Nothing
        End If
    Next
    Set Rng = Nothing
    ' Update UsedRange
    With ActiveSheet.UsedRange: End With
    GetNextLine = intCount
End Function

Private Function GetLastLine()
    GetLastLine = Range("A" & Rows.Count).End(xlUp).Row
End Function







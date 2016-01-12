VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCombinedProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'this method is prefered to other ways of opening a file because the focus changes to the file _
'and the popup boxes pop up in front of other windows
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Private pFileAddress As String
Private pCombinedProject As MSProject.Project
Public Projects As Collection
Private pRawData As Workbook
Private pCheckBox As OLEObject

' This section turns on high overhead operations
Private Sub AppTrue()
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

' This section increases performance by turning off high overhead operations
Private Sub AppFalse()
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

End Sub
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Property Get RawData() As Workbook
        Set RawData = pRawData
End Property
Public Property Let RawData(value As Workbook)
        Set pRawData = value
End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||
Private Sub Class_Initialize()
    Set Projects = New Collection
End Sub
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Property Get CombinedProject() As MSProject.Project
        Set CombinedProject = pCombinedProject
End Property
Public Property Let CombinedProject(value As MSProject.Project)
        Set pCombinedProject = value
End Property
'||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Property Get FileAddress() As String
    FileAddress = pFileAddress
End Property
Public Property Let FileAddress(value As String)
    pFileAddress = value
End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Sub GetRawData()
    
    Dim newHour As Variant
    Dim newMinute As Variant
    Dim newSecond As Variant
    Dim WaitTime As Variant
    Dim SaveFilePath As String
    Dim i As Integer
    
    Application.ActivateMicrosoftApp (xlMicrosoftProject)
    Me.CombinedProject.Views("Gantt Chart").Apply
        'select all tasks
    appProj.Visible = True
    appProj.SelectAll
        'OutlineShowAllTasks called redundantly because that seems to sometimes be needed in master schedules.  May be redundant in VBA.
    appProj.OutlineShowAllTasks
    appProj.OutlineShowAllTasks
    appProj.OutlineShowAllTasks
        'clear autofilters
    appProj.FilterClear
        'calculate
    appProj.CalculateAll
        'apply metlite view
    Me.CombinedProject.Views("Combined Met Lite").Apply
    appProj.SelectAll
    SendKeys "^c"
        'Return focus to Excel
    AppActivate "Microsoft Excel"
        'Application.wait added because SendKeys was not working without it
    newHour = Hour(Now())
    newMinute = Minute(Now())
    newSecond = Second(Now()) + 5
    WaitTime = TimeSerial(newHour, newMinute, newSecond)
    Application.Wait WaitTime
        'add new workbook for combined raw data
    Workbooks.Add
    Me.RawData = ActiveWorkbook
        'paste data into excel
    Me.RawData.Sheets(1).Paste
        'close project, don't save
    appProj.FileExit (pjDoNotSave)
        'move headers from sheet1 to sheet2.  Not used again, but may be helpful if user needs to open raw data file
    Me.RawData.Sheets(1).Rows(1).Copy
    Me.RawData.Sheets(2).Paste
    Me.RawData.Sheets(1).Rows(1).Delete
        'delete external links
    For i = 1 To FindLastRow
        If Me.RawData.Sheets(1).Cells(i, 1).Font.Color = RGB(128, 128, 128) Then
            Me.RawData.Sheets(1).Rows(i).Delete
            i = i - 1
        End If
    Next i
    
        'define save file path
    SaveFilePath = MetLitePath & "\Combined" & "\" & StatusYear & "_" & StatusMonth & "_" & StatusDay
    
        'Check if file already exists and delete it if it does
    If Dir(SaveFilePath & "*") <> vbNullString Then
    On Error Resume Next
    Workbooks(StatusYear & "_" & StatusMonth & "_" & StatusDay & ".xlsx").Close SaveChanges:=False
    On Error GoTo 0
        Kill (SaveFilePath & "*")
    End If
    
        'create folder if it does not exist and save combined file
    If Not FS.FolderExists(MetLitePath) Then
        MkDir (MetLitePath)
    End If
    If FS.FolderExists(MetLitePath & "\Combined") Then
        Me.RawData.SaveAs Filename:=SaveFilePath
    ElseIf FS.FolderExists(MetLitePath) Then
        MkDir (MetLitePath & "\Combined")
        Me.RawData.SaveAs Filename:=SaveFilePath
    End If
    
End Sub

Public Sub SplitRawData()
        'break out individual Raw data files to their own workbooks and save
    Dim SchedStart As Long
    Dim SchedFinish As Long
    Dim Project As CProject
    Dim SaveFolder As String
    Dim SaveFilePath As String
    Dim i As Integer
    Dim j As Integer
    Dim NumRows As Integer
    
    Me.RawData.Activate
    
    j = 1
    NumRows = FindLastRow
    'for eaach row in the combined raw data file
    For i = 1 To NumRows
            'if a cell in column A does not have anything in front of the "-" then set to schedstart
            If Left(Cells(i, 1), 1) = "-" Then
                SchedStart = i + 1
            End If
            'if a cell in column A (starting at row 2) does not have anything in front of the "-" then set to schedfinish
            If Left(Cells(i + 1, 1), 1) = "-" Or Cells(i + 1, 1) = "" Then
                SchedFinish = i
                'if that schedules checkbox is checked
                If Me.Projects.Item(j).CheckBox.Object.value = True Then
                    'select, group, copy, add to new workbook, save
                    Range(Cells(SchedStart, 1), Cells(SchedFinish, 24)).Select
                    Selection.Rows.Group
                    Selection.Copy
                    Workbooks.Add
                    ActiveWorkbook.Sheets(1).Paste
                    Me.Projects.Item(j).RawData = ActiveWorkbook
                    SaveFolder = Left(Cells(1, 1), InStrRev(Cells(1, 1), "-") - 1)
                    SaveFilePath = MetLitePath & "\" & SaveFolder & "\" & StatusMonth & "_" & StatusDay & "_" & StatusYear
                    'if the file exists delete it
                    If Dir(SaveFilePath & "*") <> vbNullString Then
                        On Error Resume Next
                        Workbooks(StatusMonth & "_" & StatusDay & "_" & StatusYear & ".xlsx").Close SaveChanges:=False
                        On Error GoTo 0
                        Kill (SaveFilePath & "*")
                    End If
                    
                    'check if project folder exists.  If not create it. Either way, save the raw data file in the project folder
                    If FS.FolderExists(MetLitePath & "\" & SaveFolder) Then
                        Me.Projects.Item(j).RawData.SaveAs Filename:=SaveFilePath
                        Me.Projects.Item(j).RawData.Close
                    ElseIf FS.FolderExists(MetLitePath) Then
                        MkDir (MetLitePath & "\" & SaveFolder)
                        Me.Projects.Item(j).RawData.SaveAs Filename:=SaveFilePath
                        Me.Projects.Item(j).RawData.Close
                    End If
                    'if the checkbox is not checked, delete the rows for that project and adjust the value of i by the number of deleted rows.
                ElseIf Me.Projects.Item(j).CheckBox.Object.value = False Then
                    Range(Cells(SchedStart - 1, 1), Cells(SchedFinish, 24)).Select
                    i = i - Selection.Rows.Count
                    NumRows = NumRows - Selection.Rows.Count
                    Selection.Rows.Delete
                End If
                'iterate j to move to the next checkbox
                If j < Me.Projects.Count Then
                    j = j + 1
                'exit for if j = thenumbe of schedules
                ElseIf j = Me.Projects.Count Then
                    Exit For
                End If
            End If
    Next i
    'close raw data and save
    Me.RawData.Close (True)
End Sub
    'use the API function (at the top) to open the file and set to the combinedproject property
    'using this function is the best method of opening the file because it gives the window focus and displays dialog boxes as if the user opened the file
Public Sub FileOpen()
        ShellExecute 0, vbNullString, FileAddress, 0&, 0&, 1
        Set pCombinedProject = appProj.ActiveProject
End Sub
    'save the file and assign the fileaddress property
    Public Sub Save(Address)
        CombinedProject.SaveAs Address, pjMPP
        pFileAddress = Address & ".mpp"
    End Sub
    'close and do not save
Public Sub FileClose()
        Me.CombinedProject.Activate
        appProj.FileCloseEx pjDoNotSave
End Sub
    Public Sub GetSubProjects(Optional ByVal FromExcel As Boolean = False)
        
        Dim ProjectName As String
        Dim task As task
        Dim Project As CProject
        Dim i As Integer
        
        If FromExcel = False Then
            'for each task in combined
            For Each task In Me.CombinedProject.Tasks
                'put the paths of the subprojects in SchedulePaths Array
                If task.Project Like "*Combined*" Then
                    'fill the projects collection and assign the Name and Fileaddress properties for each project
                    ProjectName = RightOf(task.Subproject, "\")
                    ProjectName = LeftOf(ProjectName, ".")
                    Set Project = New CProject
                    Project.Name = ProjectName
                    Project.FileAddress = task.Subproject
                    Me.Projects.Add Project
                End If
            Next task
            
        ElseIf FromExcel = True Then
            'for the entry area on the MacroWB
            For i = 2 To 30
                'if there is a schedule name in the cell
                If MacroWB.Sheets(1).Cells(i, 1) <> "" Then
                    'fill the projects collection and assign the Text8 and MetLitePath properties
                    Set Project = New CProject
                    Project.Text8 = MacroWB.Sheets(1).Cells(i, 1)
                    Project.MetLitePath = MetLitePath & "\" & Project.Text8 & "\" & Project.Text8 & " " & MacroWB.Sheets(1).Cells(2, 5)
                    Me.Projects.Add Project
                End If
            Next i
        End If
                
    End Sub
Public Sub CheckText8()
    
    Dim Project As CProject
    Dim Count As Integer
    Dim task As task
    Dim InvalidMsg As String
    Dim EmptyMsg As String
    Dim SaveIt As Boolean
    SaveIt = False
    
    For Each Project In Me.Projects
        Project.FileOpen
        'default the invalidtext8 and text8notfound properties
        Project.InvalidText8 = False
        Project.Text8NotFound = True
        Count = 0
        
        For Each task In Project.Project.Tasks
            'check the first 5 tasks for a text 8 entry
            If Count < 5 Then
                Count = Count + 1
                'check for invalid text8 characters
                If InStr(1, task.Text8, "\") > 0 Or InStr(1, task.Text8, "/") > 0 Or InStr(1, task.Text8, ":") > 0 Or InStr(1, task.Text8, "*") > 0 _
                    Or InStr(1, task.Text8, "?") > 0 Or InStr(1, task.Text8, Chr(34)) > 0 Or InStr(1, task.Text8, "<") > 0 _
                    Or InStr(1, task.Text8, ">") > 0 Or InStr(1, task.Text8, "|") > 0 Or InStr(1, task.Text8, ".") > 0 Then
                        Project.InvalidText8 = True
                End If
                'set text8 and text8notfound
                If task.Text8 <> "" Then
                    Project.Text8NotFound = False
                    Project.Text8 = task.Text8
                End If
            Else
                'if Count is not < 5
                Exit For
            End If
        Next task
        
        If Project.InvalidText8 = True Or Project.Text8NotFound = True Then
            Project.FileClose
        Else
            'fill in any blanks with the found text8 value
            For Each task In Project.Project.Tasks
                If task.Text8 = "" Then
                    task.Text8 = Project.Text8
                    SaveIt = True
                End If
            Next task
            'if a text8 was filled in than save the project when it is closed
            Project.FileClose (SaveIt)
        End If
    Next Project
    
    InvalidMsg = ""
    EmptyMsg = ""
    'set up msgbox message
    For Each Project In Me.Projects
        If Project.InvalidText8 = True Then
            If InvalidMsg = "" Then
                InvalidMsg = Project.Name
            Else
                InvalidMsg = InvalidMsg & ", " & Project.Name
            End If
        End If
        'set up msgbox message
        If Project.Text8NotFound = True Then
            If EmptyMsg = "" Then
                EmptyMsg = Project.Name
            Else
                EmptyMsg = EmptyMsg & ", " & Project.Name
            End If
        End If
    Next Project
    'if msgbox message strings are not still blank display msgbox
    If InvalidMsg <> "" Then
        AppActivate "Microsoft Excel"
        MacroWB.Activate
        MsgBox "Invalid characters have been found for the following schedules in the Text 8 field: " & vbCrLf & vbCrLf & InvalidMsg & vbCrLf & vbCrLf & "Correct these errors and retry"
    End If
    'if msgbox message strings are not still blank display msgbox
    If EmptyMsg <> "" Then
        AppActivate "Microsoft Excel"
        MacroWB.Activate
        MsgBox "Text8 appears to be empty for the following schedules:" & vbCrLf & vbCrLf & EmptyMsg & vbCrLf & vbCrLf & "Correct these errors and retry"
    End If
    'if msgbox message strings are not still blank display msgbox
    If EmptyMsg = "" And InvalidMsg = "" Then
        AppActivate "Microsoft Excel"
        MacroWB.Activate
        MsgBox "Check Text8 Complete, No errors found"
    End If
    'activate project
    Application.ActivateMicrosoftApp (xlMicrosoftProject)
    
End Sub

Public Sub CreateCheckBoxes()

    Dim Project As CProject
    Dim Top As Variant
    Dim i As Integer
    Dim OLEObject As OLEObject
    'delete checkboxes
    For Each OLEObject In MacroWB.Sheets(1).OLEObjects
        If InStr(OLEObject.Name, "CheckBox") > 0 Then
            OLEObject.Delete
        End If
    Next OLEObject
    'create Combiend Checkbox
    Set pCheckBox = MacroWB.Sheets(1).OLEObjects.Add("Forms.CheckBox.1")
    pCheckBox.Top = 23.25
    'height and width based upon cell size in an attempt to make itlook the same on other people's screens
    pCheckBox.Height = MacroWB.Sheets(1).Range("A2").Height - 1.25 '21.25
    pCheckBox.Left = 1.5
    pCheckBox.Width = MacroWB.Sheets(1).Range("A2").Width - 2.75 '166.75
    pCheckBox.Name = "CheckBox1"
    pCheckBox.Object.Caption = Cells(2, 1)
    pCheckBox.Object.BackColor = &H8000000F
    pCheckBox.Object.Alignment = 1
    pCheckBox.Object.TextAlign = 1
    pCheckBox.Object.value = True
    
    Top = 45.5
    i = 3
    'create checkbox for each project
    For Each Project In Me.Projects
        Project.CheckBox = MacroWB.Sheets(1).OLEObjects.Add("Forms.CheckBox.1")
        Project.CheckBox.Top = Top
        'height and width based upon cell size in an attempt to make itlook the same on other people's screens
        Project.CheckBox.Height = MacroWB.Sheets(1).Range("A2").Height - 1.25 '21.25
        Project.CheckBox.Left = 1.5
        Project.CheckBox.Width = MacroWB.Sheets(1).Range("A2").Width - 2.75 '166.75
        Project.CheckBox.Name = "CheckBox" & i - 1
        Project.CheckBox.Object.Caption = Cells(i, 1)
        Project.CheckBox.Object.BackColor = &H8000000F
        Project.CheckBox.Object.Alignment = 1
        Project.CheckBox.Object.TextAlign = 1
        Project.CheckBox.Object.value = True
        
        Top = Top + MacroWB.Sheets(1).Range("A2").Height '22.5
        i = i + 1
        
    Next Project
    
End Sub

Public Sub RunMetLite()

    Dim Project As CProject
    Dim FirstTimeRun As Boolean
    Dim j As Integer
    Dim k As Integer
    Dim Duplicate As Integer
    Dim CEIDataName As String
    Dim CEIPath As String
    Dim ActionableItemsName As String
 
    For Each Project In Me.Projects
        'if checkbox is checked
        If Project.CheckBox.Object.value = True Then
            Project.MetLitePath = MetLitePath & "\" & Project.Text8 & "\" & Project.Text8 & " " & MacroWB.Sheets(1).Cells(2, 5)
            'if the file is there
            If FS.FileExists(Project.MetLitePath) Then
                Workbooks.Open (Project.MetLitePath)
                Project.MetLite = ActiveWorkbook
                Project.MetLite.Sheets("Input").Activate
                    'fill in the filepath and program name
                    Cells(7, 7) = MetLitePath & "\" & Project.Text8 & "\"
                    Cells(2, 2) = Project.Text8
                    'find the next open spot
                    For j = 82 To 11 Step -1
                        If Cells(j, 7) <> "" Then
                            'fill in settings for the new row, following different naming convention for combined
                            If Project.Text8 = "Combined" Then
                                Cells(j + 1, 7) = StatusYear & "_" & StatusMonth & "_" & StatusDay
                            Else
                                Cells(j + 1, 7) = StatusMonth & "_" & StatusDay & "_" & StatusYear
                            End If
                            Cells(35, 2) = Cells(j + 1, 7)
                            Cells(j + 1, 6) = StatusMonth & "/01/" & StatusYear
                            Cells(j + 1, 5) = StatusDate & " 17:00"
                            Cells(j + 1, 4) = "Yes"
                            'make sure the previous months settings are set to 'No'
                            For k = j To 11 Step -1
                                Cells(k, 4) = "No"
                                Cells(k, 12) = "No"
                                Cells(k, 13) = "No"
                            Next
                            'finish filling out current month settings
                            Cells(j + 1, 8) = 12
                            Cells(j + 1, 12) = "Yes"
                            Cells(j + 1, 13) = "No"
                            Cells(4, 17) = "No"
                            Exit For
                        End If
                    Next j
                    'if this is the first time running through MetLite
                    If Cells(11, 7) = "" Then
                        'fill in settings for the new row, following different naming convention for combined
                        If Project.Text8 = "Combined" Then
                            Cells(11, 7) = StatusYear & "_" & StatusMonth & "_" & StatusDay
                        Else
                            Cells(11, 7) = StatusMonth & "_" & StatusDay & "_" & StatusYear
                        End If
                        Cells(35, 2) = Cells(11, 7)
                        Cells(11, 6) = StatusMonth & "/01/" & StatusYear
                        Cells(11, 5) = StatusDate & " 17:00"
                        Cells(11, 4) = "Yes"
                        'finish filling out current month settings
                        Cells(11, 8) = 12
                        Cells(11, 12) = "Yes"
                        Cells(11, 13) = "No"
                        Cells(4, 17) = "No"
                        FirstTimeRun = True
                        
                    End If
                    
                    Project.MetLite.Sheets("Input").Activate
                    Cells(j + 1, 7).Select
    
                    'check if re-run and copy over previous run data if true
                    If Cells(j + 1, 7) = Cells(j, 7) Then
                        Range(Cells(j + 1, 4), Cells(j + 1, 13)).Copy
                        Range(Cells(j, 4), Cells(j, 13)).Select
                        ActiveSheet.Paste
                        Range(Cells(j + 1, 4), Cells(j + 1, 13)) = ""
                    End If
                    
                    'Run Metlite
                    Call AppTrue
                    Application.Run ("'" & Project.MetLite.Name & "'!main")
                    Project.MetLite.Sheets("Input").Activate
                    'find the last entry
                    j = 82
                    While Cells(j, 7) = ""
                        j = j - 1
                    Wend
                    'set up options to run CEI
                    If j > 11 Then
                        Cells(j, 4) = "Yes"
                        Cells(j - 1, 4) = "Yes"
                        Cells(j, 12) = "No"
                    'capture which schedules do not have 2 months data
                    ElseIf j <= 11 Then
                        Project.NotTwoMonths = True
                    End If
                    'capture schedules that have two months data but the past data file is missing
                    If FS.FileExists(Project.MetLite.Sheets("Input").Cells(7, 7) & Project.MetLite.Sheets("Input").Cells(j - 1, 7) & ".xlsx") Then
                        Application.Run ("'" & Project.MetLite.Name & "'!CEI")
                        'find CEIData filename
                        CEIDataName = Dir(MetLitePath & "\" & Project.Text8 & "\" & "*CEI DATA*")
                        CEIPath = MetLitePath & "\" & Project.Text8 & "\" & CEIDataName
                        Workbooks.Open (CEIPath)
                        Project.CEI = ActiveWorkbook
                        Project.CEI.Close (True)
                    ElseIf Not FS.FileExists(Project.MetLite.Sheets("Input").Cells(7, 7) & Project.MetLite.Sheets("Input").Cells(j - 1, 7)) Then
                        Project.LastMonthMissing = True
                    End If
                       
                    Project.MetLite.Close (True)
                    Call AppTrue

    
                    'find actionable item file name
                    ActionableItemsName = Dir(MetLitePath & "\" & Project.Text8 & "\" & "*Actionable Schedule*")
                    'create Year folder if needed
                    If Not FS.FolderExists(MetLitePath & "\" & Project.Text8 & "\" & StatusYear) Then
                        MkDir (MetLitePath & "\" & Project.Text8 & "\" & StatusYear)
                    End If
                    'create month folder if needed
                    If Not FS.FolderExists(MetLitePath & "\" & Project.Text8 & "\" & StatusYear & "\" & StatusMonth) Then
                        MkDir (MetLitePath & "\" & Project.Text8 & "\" & StatusYear & "\" & StatusMonth)
                    End If
                    'if actionable items file already exists delete the old one
                    If FS.FileExists(MetLitePath & "\" & Project.Text8 & "\" & StatusYear & "\" & StatusMonth & "\" & ActionableItemsName) Then
                        Kill (MetLitePath & "\" & Project.Text8 & "\" & StatusYear & "\" & StatusMonth & "\" & ActionableItemsName)
                    End If
                    'if CEI DATA file already exists delete the old one
                    If FS.FileExists(MetLitePath & "\" & Project.Text8 & "\" & StatusYear & "\" & StatusMonth & "\" & CEIDataName) Then
                        Kill (MetLitePath & "\" & Project.Text8 & "\" & StatusYear & "\" & StatusMonth & "\" & CEIDataName)
                    End If
                    
                    'move the Actionable Items file from where MetLite saved it to where it should be in the file structure
                    If FS.FileExists(MetLitePath & "\" & Project.Text8 & "\" & ActionableItemsName) Then
                        Name MetLitePath & "\" & Project.Text8 & "\" & ActionableItemsName As MetLitePath & "\" & Project.Text8 & "\" & StatusYear & "\" & StatusMonth & "\" & ActionableItemsName
                    End If
                    'move the CEI DATA file from where MetLite saved it to where it should be in the file structure
                    If FS.FileExists(MetLitePath & "\" & Project.Text8 & "\" & CEIDataName) Then
                        Name MetLitePath & "\" & Project.Text8 & "\" & CEIDataName As MetLitePath & "\" & Project.Text8 & "\" & StatusYear & "\" & StatusMonth & "\" & CEIDataName
                    End If
                    'If metlite doesn't exist then add schedule name to a missing metlite array
                ElseIf Not FS.FileExists(MetLitePath & "\" & Project.Text8 & "\" & Project.Text8 & " " & Cells(2, 5)) Then
                    Project.MissingMetLite = True
                End If
            End If
        Next Project
End Sub
Public Sub GetCheckBoxes()
    
    Dim i As Integer
    i = 1
    Dim OLEObject As OLEObject
    Dim Project As CProject
    'for each project go through the checkboxes and find one qith a caption that matches the projects text8 property
    For Each Project In Me.Projects
        For Each OLEObject In MacroWB.Sheets(1).OLEObjects
            If OLEObject.progID = "Forms.CheckBox.1" Then
                If OLEObject.Object.Caption = Project.Text8 Then
                    Project.CheckBox = OLEObject
                End If
            End If
        Next OLEObject
    Next Project
    
End Sub
    'grab the string tothe right of the last instance of "Symbol"
    Function RightOf(SourceString, Symbol)
        RightOf = Right(SourceString, Len(SourceString) - InStrRev(SourceString, Symbol))
    End Function
    'grab the string to the left of the first instance of "Symbol"
    Function LeftOf(SourceString, Symbol)
        LeftOf = Left(SourceString, Len(SourceString) - (Len(SourceString) - (InStr(SourceString, Symbol) - 1)))
    End Function
    
    ' Find last row on sheet
Function FindLastRow() As Integer
    
    If WorksheetFunction.CountA(Cells) > 0 Then
        FindLastRow = Cells.Find(What:="*", After:=[A1], _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    End If
         
End Function


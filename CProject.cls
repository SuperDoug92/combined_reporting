VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CProject"
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
Private pName As String
Private pNumMissingStatus As Integer
Private pNumConstraints As Integer
Private pProject As MSProject.Project
Private pMissingStatusUID() As String
Private pConstraintAddedUID() As String
Private pInvalidText8 As Boolean
Private pText8NotFound As Boolean
Private pText8 As String
Private pRawData As Workbook
Private pCheckBox As OLEObject
Private pID As Integer
Private pMetLite As Workbook
Private pMetLitePath As String
Private pMissingMetLite As Boolean
Private pNotTwoMonths As Boolean
Private pLastMonthMissing As Boolean
Private pCEI As Workbook
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    Public Property Get CEI() As Workbook
        Set CEI = pCEI
    End Property
    Public Property Let CEI(value As Workbook)
        Set pCEI = value
    End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Property Get LastMonthMissing() As Boolean
        LastMonthMissing = pLastMonthMissing
End Property
Public Property Let LastMonthMissing(value As Boolean)
        pLastMonthMissing = value
End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    Public Property Get NotTwoMonths() As Boolean
        NotTwoMonths = pNotTwoMonths
    End Property
    Public Property Let NotTwoMonths(value As Boolean)
        pNotTwoMonths = value
    End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Property Get MissingMetLite() As Boolean
        MissingMetLite = pMissingMetLite
End Property
Public Property Let MissingMetLite(value As Boolean)
        pMissingMetLite = value
End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    Public Property Get MetLitePath() As String
        MetLitePath = pMetLitePath
    End Property
    Public Property Let MetLitePath(value As String)
        pMetLitePath = value
    End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Property Get MetLite() As Workbook
        Set MetLite = pMetLite
End Property
Public Property Let MetLite(value As Workbook)
        Set pMetLite = value
End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    Public Property Get ID() As Integer
        ID = pID
    End Property
    Public Property Let ID(value As Integer)
        pID = value
    End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Property Let CheckBox(value As OLEObject)
        Set pCheckBox = value
End Property
Public Property Get CheckBox() As OLEObject
        Set CheckBox = pCheckBox
End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    Public Property Get RawData() As Workbook
        Set RawData = pRawData
    End Property
    Public Property Let RawData(value As Workbook)
        Set pRawData = value
    End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Property Get Text8() As String
        Text8 = pText8
End Property
Public Property Let Text8(value As String)
        pText8 = value
End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    Public Property Get InvalidText8() As Boolean
        InvalidText8 = pInvalidText8
    End Property
    Public Property Let InvalidText8(value As Boolean)
        pInvalidText8 = value
    End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Property Get Text8NotFound() As Boolean
        Text8NotFound = pText8NotFound
End Property
Public Property Let Text8NotFound(value As Boolean)
        pText8NotFound = value
End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    Public Property Get MissingStatusUID() As Variant
        MissingStatusUID = pMissingStatusUID
    End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Property Get ConstraintAddedUID() As MSProject.Project
        ConstraintAddedUID() = pConstraintAddedUID()
End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    Public Property Get NumMissingStatus() As Integer
        NumMissingStatus = pNumMissingStatus
    End Property
    Public Property Let NumMissingStatus(value As Integer)
        pNumMissingStatus = value
    End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Property Get NumConstraints() As Integer
        NumConstraints = pNumConstraints
End Property
Public Property Let NumConstraints(value As Integer)
        pNumConstraints = value
End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    Public Property Get FileAddress() As String
        FileAddress = pFileAddress
    End Property
    Public Property Let FileAddress(value As String)
        pFileAddress = value
    End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
Public Property Get Name() As String
        Name = pName
End Property
Public Property Let Name(value As String)
        pName = value
End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    Public Property Let Project(value As MSProject.Project)
        Set pProject = value
    End Property
    Public Property Get Project() As MSProject.Project
        Set Project = pProject
    End Property
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||
    Public Sub FileOpen()
        ShellExecute 0, vbNullString, FileAddress, 0&, 0&, 1
        Set pProject = appProj.ActiveProject
    End Sub
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||

Public Sub Status()

        Dim LWDStatus As Date
        Dim j As Integer
        Dim ForecastStart As Variant
        Dim task As task
        Dim Filtername As String
        Dim x As Integer
        x = 1
        
        
        'in the month and year of the status date
        With Me.Project.Calendar.Years(StatusYear).Months(CInt(StatusMonth))
            'counting backwards from the last day of the month
            For j = .Days.Count To 1 Step -1
                'if it is a working day then assign that day to LWDStatus variable
                If .Days(j).Working = True Then
                    LWDStatus = StatusMonth & "/" & j & "/" & StatusYear & " 17:00"
                    Exit For
                End If
            Next j
        End With
        
        appProj.FilterClear
        appProj.OutlineShowAllTasks
        'if the user has supplied a filter then apply the filter or report its failing to the user
        If MacroWB.Sheets(1).Cells(5, 3) <> "" Then
            Filtername = MacroWB.Sheets(1).Cells(5, 3)
            On Error GoTo FilterDNE
            Me.Project.TaskFilters(Filtername).Apply
            If x = 2 Then
FilterDNE:      MsgBox "The filter does not exist and could not be applied in" & Project.Text8 & vbCrLf & "Improper status will be checked for all tasks"
            End If
            appProj.SelectAll
            
            On Error GoTo 0
            
        End If
        
        Me.Project.AutoSplitTasks = True
        Me.Project.StatusDate = StatusDate & " 17:00"
        'back up status to the status date
        For Each task In ActiveSelection.Tasks
            ForecastStart = task.Start
            If task.Stop > Me.Project.StatusDate And task.Stop <> "NA" And task.ExternalTask <> True And Not task.Summary Then
                If ForecastStart <= Me.Project.StatusDate Then
                    task.Stop = Me.Project.StatusDate
                ElseIf ForecastStart >= Me.Project.StatusDate Then
                    task.Stop = "NA"
                End If
                appProj.CalculateProject
                'add constraints and capture which tasks by UID
                If ForecastStart > task.Start Then
                    NumConstraints = NumConstraints + 1
                    task.ConstraintType = "Start No Earlier Than"
                    task.ConstraintDate = ForecastStart
'                    ReDim Preserve ConstraintAddedProject(NumConstraints)
'                    ConstraintAddedProject(NumConstraints) = Task.Project
                    ReDim Preserve pConstraintAddedUID(pNumConstraints)
                    pConstraintAddedUID(pNumConstraints) = task.UniqueID
'                    ReDim Preserve ConstraintAddedName(NumConstraints)
'                    ConstraintAddedName(NumConstraints) = Task.Name
                End If
            End If
            'capture missing status tasks by UID
            If task.Stop < LWDStatus And task.ExternalTask <> True And task.ActualFinish = "NA" And Not task.Summary Then
                NumMissingStatus = NumMissingStatus + 1
'                ReDim Preserve MissingStatusProject(NumMissingStatus)
'                MissingStatusProject(NumMissingStatus) = Task.Project
                ReDim Preserve pMissingStatusUID(pNumMissingStatus)
                pMissingStatusUID(pNumMissingStatus) = task.UniqueID
'                ReDim Preserve MissingStatusName(NumMissingStatus)
'                MissingStatusName(NumMissingStatus) = Task.Name
                'ReDim Preserve StopDate(NumMissingStatus)
                'StopDate(NumMissingStatus) = Task.Stop
            End If
            'capture missing status tasks by UID
            If task.Stop = "NA" And task.Start <= LWDStatus And task.ExternalTask <> True And Not task.Summary Then
                NumMissingStatus = NumMissingStatus + 1
'                ReDim Preserve MissingStatusProject(NumMissingStatus)
'                MissingStatusProject(NumMissingStatus) = Task.Project
                ReDim Preserve pMissingStatusUID(pNumMissingStatus)
                pMissingStatusUID(pNumMissingStatus) = task.UniqueID
'                ReDim Preserve MissingStatusName(NumMissingStatus)
'                MissingStatusName(NumMissingStatus) = Task.Name
                'ReDim Preserve StopDate(NumMissingStatus)
                'StopDate(NumMissingStatus) = Task.Stop
            End If
        Next task
        'turn off Autosplit tasks
        Me.Project.AutoSplitTasks = False
        
End Sub
'delete file if it exists andsave to the monthly file location
Public Sub SaveAsMonthly()
        If FS.FileExists(EOMSchedulePath & "\" & LeftOf(Me.Project.Name, ".") & " " & StatusMonth & "-" & StatusDay & "-" & StatusYear & ".mpp") Then
            Kill EOMSchedulePath & "\" & LeftOf(Me.Project.Name, ".") & " " & StatusMonth & "-" & StatusDay & "-" & StatusYear & ".mpp"
        End If
        
        Me.Project.SaveAs Name:=EOMSchedulePath & "\" & LeftOfLast(Me.Project.Name, ".") & " " & StatusMonth & "-" & StatusDay & "-" & StatusYear, Format:=pjMPP
        
End Sub
    ' close and save if user passes True
    Public Sub FileClose(Optional ByVal Save As Boolean = False)
        Me.Project.Activate
        If Save = False Then
            appProj.FileCloseEx pjDoNotSave
        ElseIf Save = True Then
            appProj.FileCloseEx pjSave
        End If
        
    End Sub

Function RightOf(SourceString, Symbol)
        RightOf = Right(SourceString, Len(SourceString) - InStrRev(SourceString, Symbol))
End Function
Function LeftOf(SourceString, Symbol)
        LeftOf = Left(SourceString, Len(SourceString) - (Len(SourceString) - (InStr(SourceString, Symbol) - 1)))
End Function
'left of last instance of "Symbol"
Function LeftOfLast(SourceString, Symbol)
    LeftOfLast = Left(SourceString, Len(SourceString) - (Len(SourceString) - (InStrRev(SourceString, Symbol) - 1)))
End Function

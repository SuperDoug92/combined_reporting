Attribute VB_Name = "Common"
Option Explicit

'this method is prefered to other ways of opening a file because the focus changes to the file _
'and the popup boxes pop up in front of other windows
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' This section turns on high overhead operations
Sub AppTrue()
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub

' This section increases performance by turning off high overhead operations
Sub AppFalse()
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

End Sub

'This section turns off MSProject overhead operations
Sub ProjAppFalse()
    appProj.ScreenUpdating = False
    appProj.DisplayAlerts = False
    
End Sub

'This section turns off MSProject overhead operations
Sub ProjAppTrue()
    appProj.ScreenUpdating = True
    appProj.DisplayAlerts = True

End Sub
'a sub that can be called stand alone to test new ideas
Sub Test()
    Call SetUp
    Dim CombinedName As String
    Dim CurrentCombined As New CCombinedProject
    Dim appProj As MSProject.Application
    Set appProj = CreateObject("MSProject.Application")
    Dim value As Double
    Dim ShellString As String
    
    'Set MS Project Object, Open Combined, Set combined object
    CombinedName = Dir(SchedulePath & "\" & "*Combined*")
    CurrentCombined.FileAddress = SchedulePath & "\" & CombinedName
    CurrentCombined.FileOpen
    appProj.Visible = False
    'ActiveWindow.Application.FileOpenEx CurrentCombined.FileAddress
    'Windows(va).Application.FileOpenEx CurrentCombined.FileAddress
    'C:\Program Files (x86)\Microsoft Office\Office14\WINPROJ.exe
    
    CurrentCombined.FileAddress = SchedulePath & "\" & CombinedName
    
    appProj.FileOpenEx CurrentCombined.FileAddress
    CurrentCombined.FileOpen
    CurrentCombined.GetSubProjects
    appProj.FileExit pjDoNotSave
    CurrentCombined.CreateCheckBoxes

End Sub
    'checks text8 and fills in contract field if missing
Sub Check()
    Call AppFalse

    Dim CombinedName As String
    Dim Project As CProject
    Dim i As Integer
    
    Dim CurrentCombined As New CCombinedProject
    'Set MS Project Object, Open Combined, Set combined object
    CombinedName = Dir(SchedulePath & "\" & "*Combined*")
    CurrentCombined.FileAddress = SchedulePath & "\" & CombinedName
    CurrentCombined.FileOpen
    
    'get subprojects, check and repair text8 when possible. msg when not
    CurrentCombined.GetSubProjects
    CurrentCombined.FileClose
    CurrentCombined.CheckText8
    'close MS Project
    appProj.FileExit
    
    'unlock sheet, fill in the project text 8 values in the left columns, create checkboxes, lock shett
    MacroWB.Sheets(1).Unprotect Password:="air42"
    
    For i = 2 To 30
        MacroWB.Sheets(1).Cells(i, 1) = ""
    Next i
    
    MacroWB.Sheets(1).Cells(1, 1) = "ActiveProjects (" & StatusDate & "):"
    MacroWB.Sheets(1).Cells(2, 1) = "Combined"
    
    i = 3
    For Each Project In CurrentCombined.Projects
        Cells(i, 1) = Project.Text8
        i = i + 1
    Next Project
    
    CurrentCombined.CreateCheckBoxes
    
    MacroWB.Sheets(1).Protect Password:="air42"
    
    'reactivate background stuff
    Call AppTrue
    
End Sub

'Sets all of the schedules to the date input by the user on macro input sheet
Sub Status()

    Call AppFalse
    
    Dim CurrentCombined As New CCombinedProject
    Dim MonthlyCombined As New CCombinedProject
    
        'variable used to create the combined project file path
    Dim CombinedName As String
    
        'variables to hold the msgbox message strings
    Dim ConstraintMsgBox As String
    Dim MissingStatusMsgBox As String
    
        'collection variables for for each loops
    Dim task As task
    Dim Project As CProject
    
        'variable used in correcting the links to subprojects in the end of month combined
    Dim SubProjName As String
     
        'variables used in the creation of the message boxes
    Dim ConstraintCheck As Boolean
    Dim StatusCheck As Boolean
        
        'Set MS Project Object, Open Combined, Set combined object
    CombinedName = Dir(SchedulePath & "\" & "*Combined*")
    CurrentCombined.FileAddress = SchedulePath & "\" & CombinedName
    CurrentCombined.FileOpen
        'Set the Monthly Combined Object to the Current combined object, then save as a monthly version
    MonthlyCombined.CombinedProject = CurrentCombined.CombinedProject
        'delete file if it already axists
    If FS.FileExists(EOMSchedulePath & "\" & LeftOf(MonthlyCombined.CombinedProject.Name, ".") & " " & StatusMonth & "-" & StatusDay & "-" & StatusYear & ".mpp") Then
        Kill EOMSchedulePath & "\" & LeftOf(MonthlyCombined.CombinedProject.Name, ".") & " " & StatusMonth & "-" & StatusDay & "-" & StatusYear & ".mpp"
    End If
    
        'build end of month file location if it does not already exist
    If Not FS.FolderExists(SchedulePath & "\" & StatusYear) Then
        MkDir (SchedulePath & "\" & StatusYear)
    End If
    If Not FS.FolderExists(EOMSchedulePath) Then
        MkDir (EOMSchedulePath)
    End If
    
        'save monthly version of combined
    MonthlyCombined.Save (EOMSchedulePath & "\" & LeftOf(MonthlyCombined.CombinedProject.Name, ".") & " " & StatusMonth & "-" & StatusDay & "-" & StatusYear)
        'creat projects collection with file addresses
    Call ProjAppFalse
    MonthlyCombined.GetSubProjects
    MonthlyCombined.FileClose
    Call ProjAppTrue
    
        'open, status, save monthly version, and close each subproject file
    For Each Project In MonthlyCombined.Projects
        Project.FileOpen
        Project.Status
        Project.SaveAsMonthly
        Project.FileClose
    Next Project
            
        'fix the suproject links in the end of month combined file
        'had to rig the next line because the macro would exit the for loop after it changed the task.subproject field on the first one
        MonthlyCombined.FileOpen
reLoop: For Each task In MonthlyCombined.CombinedProject.Tasks
        
            If task.Project Like "*Combined*" Then
                If InStr(1, task.Subproject, StatusYear & "\" & StatusMonth) < 5 Then
                    SubProjName = RightOf(task.Subproject, "\")
                    task.Subproject = Left(task.Subproject, Len(task.Subproject) - Len(SubProjName)) & StatusYear & "\" & StatusMonth & "\" & Left(SubProjName, Len(SubProjName) - 4) & " " & StatusMonth & "-" & StatusDay & "-" & StatusYear & ".mpp"
                    GoTo reLoop
                End If
            End If
        Next task
        'Save monthly combined and close project
    MonthlyCombined.Save (MonthlyCombined.FileAddress)
    appProj.FileCloseAllEx pjDoNotSave
    appProj.FileExit

    ConstraintCheck = False
    StatusCheck = False
        'fill in msgbox variables
    For Each Project In MonthlyCombined.Projects
        
        If Project.NumConstraints > 0 Then
            If ConstraintCheck = False Then
                ConstraintCheck = True
                ConstraintMsgBox = Project.Name & "  " & Project.NumConstraints & "  UID's:" & vbCrLf & vbCrLf & Join(Project.ConstraintAddedUID, ", ")
            Else
                ConstraintMsgBox = ConstraintMsgBox & vbCrLf & vbCrLf & Project.Name & "  " & Project.NumConstraints & "  For the below UID's:" & vbCrLf & vbCrLf & Join(Project.ConstraintAddedUID, ", ")
            End If
        End If
        
        If Project.NumMissingStatus > 0 Then
            If StatusCheck = False Then
                StatusCheck = True
                MissingStatusMsgBox = Project.Name & "  " & Project.NumMissingStatus & "  UID's:" & vbCrLf & vbCrLf & Join(Project.MissingStatusUID, ", ")
            Else
                MissingStatusMsgBox = MissingStatusMsgBox & vbCrLf & vbCrLf & Project.Name & "  " & Project.NumMissingStatus & "  UID's:" & vbCrLf & vbCrLf & Join(Project.MissingStatusUID, ", ")
            End If
        End If
    
    Next Project
    
        'if relevant display msgboxes
    If ConstraintMsgBox <> "" Then
        MsgBox ConstraintMsgBox, , "SNET Constraints Added"
    End If
    If MissingStatusMsgBox <> "" Then
        MsgBox MissingStatusMsgBox, , "Tasks missing status"
    End If
    
        'reactivate background stuff
    Call AppTrue
    
        'report on contraints set and missing status
    
    MsgBox "Schedules statused and saved"
    
End Sub

'Generates raw data files for all of the schedules in the combined schedule file
Sub GenRawFiles()
    'deactivate background tasks
    Call AppFalse

    Dim MsgResult As String
    Dim i As Integer
    Dim CombinedName As String
    Dim MonthlyCombined As New CCombinedProject
    
        'Set MS Project Object, Open Combined, Set combined object
    CombinedName = Dir(EOMSchedulePath & "\" & "*Combined*")
    MonthlyCombined.FileAddress = EOMSchedulePath & "\" & CombinedName
    MonthlyCombined.FileOpen
    MonthlyCombined.GetSubProjects
    
    For i = 3 To MonthlyCombined.Projects.Count + 2
        MonthlyCombined.Projects.Item(i - 2).Text8 = Cells(i, 1)
    Next i
    
    MonthlyCombined.GetCheckBoxes
    MonthlyCombined.GetRawData
    MonthlyCombined.SplitRawData
    
    'If there are fewer schedules then last time delete the contents of the extra cells
    
    MsgBox "Raw Data Files Generated"
    
    Call AppTrue
End Sub

'Runs MetLite for all of the schedules which have a metlite file in the folder structure
' adding run Metlite as a method to CCombinedProject was overkill. Only redudantly called code is useful there
Sub RunMet()
    
    Dim UserMsg As String
    Dim Display As Boolean
    Dim Project As CProject
    Dim MonthlyCombined As New CCombinedProject

    MonthlyCombined.GetSubProjects (True)
    MonthlyCombined.GetCheckBoxes
    MonthlyCombined.RunMetLite
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'MsgBox outputs to user
    
    Display = False
    UserMsg = "MetLite file not found for:" & vbCrLf & vbCrLf
    
    For Each Project In MonthlyCombined.Projects
        If Project.MissingMetLite = True Then
            UserMsg = UserMsg & ", " & Project.Text8
            Display = True
        End If
    Next Project
    
    UserMsg = UserMsg & vbCrLf & vbCrLf & vbCrLf & "Previous months data no found for:" & vbCrLf & vbCrLf
    
    For Each Project In MonthlyCombined.Projects
        If Project.LastMonthMissing = True Then
            UserMsg = UserMsg & ", " & Project.Text8
            Display = True
        End If
    Next Project
    
    UserMsg = UserMsg & vbCrLf & vbCrLf & vbCrLf & "The following schedules are being run for the first time (No CEI):" & vbCrLf & vbCrLf
        
    For Each Project In MonthlyCombined.Projects
        If Project.NotTwoMonths = True Then
            UserMsg = UserMsg & Project.Text8
            Display = True
        End If
    Next Project
    
    If Display = True Then
        MsgBox UserMsg
    End If
    
    UserMsg = "MetLite/CEI run completed for:" & vbCrLf & vbCrLf
    For Each Project In MonthlyCombined.Projects
        If Project.CheckBox.Object.value = True Then
            UserMsg = UserMsg & ", " & Project.Text8
        End If
    Next Project
    
    MsgBox UserMsg
    
    'End MsgBox output to user
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       
End Sub

'Generates the schedule metrics sheet in the macro workbook

Sub GenerateScheduleMetrics()
    
    'deactivate background tasks
    Call AppFalse
    
    Dim Project As CProject
    Dim MsgResult As String
    Dim i As Integer
    Dim CombinedName As String
    Dim MonthlyCombined As New CCombinedProject
    Dim ScheduleCount As Integer
    Dim SearchFold As String
    Dim SearchFile As String
        Dim MetFileName As String
        MetFileName = MacroWB.Sheets(1).Cells(2, 5)
    Dim NumColCEI As Integer
    Dim NumCol As Integer
    Dim Worksheet As Worksheet
    Dim MissingMetFile() As String
    ReDim MissingMetFile(0)
    Dim NoCEIData As String
    Dim j As Integer
    Dim MetLite As Workbook
    
    MonthlyCombined.GetSubProjects (True)
    ScheduleCount = MonthlyCombined.Projects.Count
        
    'delete the old sheet in case of a re-run
    For Each Worksheet In MacroWB.Sheets
        If Worksheet.Name = StatusMonth & "-" & StatusDay & "-" & StatusYear & " Metrics Analysis" Then
            Sheets(StatusMonth & "-" & StatusDay & "-" & StatusYear & " Metrics Analysis").Delete
            Exit For
        End If
    Next
    
    MacroWB.Sheets("Metrics Analysis Template").Copy After:=MacroWB.Sheets("Input")
    ActiveSheet.Name = StatusMonth & "-" & StatusDay & "-" & StatusYear & " Metrics Analysis"
    
    For j = 1 To ScheduleCount - 3
        ActiveSheet.Columns(4).Copy
        Columns(4).Insert
    Next j
    
    For Each Project In MonthlyCombined.Projects
        SearchFold = MetLitePath & "\" & Project.Text8
        'Check for existance in file structure
        If FS.FolderExists(SearchFold) Then
            SearchFile = SearchFold & "\" & Project.Text8 & " " & MetFileName
            'Check for existance in file structure
            If FS.FileExists(SearchFile) Then
                'Open Schedule Metlitefile
                Workbooks.Open (SearchFile)
                Set MetLite = ActiveWorkbook
                
                MetLite.Sheets("MetricsHistory").Activate
                ' Find the last column on the metrics history sheet
                ' Over 50 columns will a higher upper value of i
                i = 0
                
                While Cells(11, 50 - i) = ""
                    i = i + 1
                Wend
                
                i = 50 - i
                NumCol = i
                If IsNumeric(Cells(11, 50 - i)) Then
                    'Find the index of the current Schedule in the Schedules() array
                    'index = Application.Match(Schedule, Schedules, False)
                    'Select and copy data from the open Metlite file
                    Range(Cells(11, NumCol), Cells(91, NumCol)).Select
                    Selection.Copy
                    MacroWB.Sheets(StatusMonth & "-" & StatusDay & "-" & StatusYear & " Metrics Analysis").Activate
                    'Account for the "Total of individuals column"
                    If Project.Text8 = "Combined" Then
                        Cells(4, ScheduleCount + 3).Select
                        ActiveSheet.Paste
                        Cells(3, ScheduleCount + 3) = Project.Text8 & " " & StatusDate
                    ElseIf Not Project.Text8 = "Combined" Then
                        For i = 1 To ScheduleCount - 1
                            If Cells(4, i + 2) = "" Then
                                Cells(4, i + 2).Select
                                ActiveSheet.Paste
                                Cells(3, i + 2) = Project.Text8 & " " & StatusDate
                                Exit For
                            End If
                        Next i
                    End If
                End If
                
                'Active the CEI Finish History sheet on the open Metlite Workbook
                MetLite.Activate
                Sheets("CEI Finish - History").Activate
                'Find the last column on the CEI Finish History Sheet
                'More than 50 columns will require a higher upper value of i
                i = 0
                While Cells(21, 50 - i) = ""
                    i = i + 1
                Wend
                
                i = 50 - i
                
                NoCEIData = ""
                If IsNumeric(Cells(21, i)) Then
                    NumColCEI = i
                    'Select and Copy CEI Data
                    Range(Cells(21, NumColCEI), Cells(22, NumColCEI)).Select
                    Selection.Copy
                Else
                    NoCEIData = "No Data"
                End If
                
                'Paste data into Schedule Metrics Analysis Workbook
                MacroWB.Sheets(StatusMonth & "-" & StatusDay & "-" & StatusYear & " Metrics Analysis").Activate
                If Project.Text8 = "Combined" Then
                    Cells(90, ScheduleCount + 3).Select
                    ActiveSheet.Paste
                    Cells(89, ScheduleCount + 3) = Project.Text8
                ElseIf Not Project.Text8 = "Combined" Then
                    For j = ScheduleCount + 3 To 3 Step -1
                        If Cells(3, j) = Project.Text8 & " " & StatusDate Then
                            If NoCEIData = "" Then
                                Cells(90, j).Select
                                ActiveSheet.Paste
                                Cells(89, j) = Project.Text8
                                Exit For
                            ElseIf NoCEIData = "No Data" Then
                                Cells(90, j) = NoCEIData
                                Cells(89, j) = Project.Text8
                                Exit For
                            End If
                        End If
                    Next j
                End If

                'Select and Copy Volatility Data
                MetLite.Activate
                Sheets("CEI Finish - History").Activate
                Cells(20, NumColCEI).Select
                Selection.Copy
                'Paste data into Schedule Metrics Analysis Workbook
                MacroWB.Sheets(StatusMonth & "-" & StatusDay & "-" & StatusYear & " Metrics Analysis").Activate
                If Project.Text8 = "Combined" Then
                    Cells(93, ScheduleCount + 3).Select
                    ActiveSheet.Paste
                ElseIf Not Project.Text8 = "Combined" Then
                    For j = ScheduleCount + 3 To 3 Step -1
                        'look through cell 90 because it will contain "No Data" if CEI was not run forthat schedule
                        If Cells(3, j) = Project.Text8 & " " & StatusDate Then
                            Cells(93, j).Select
                            ActiveSheet.Paste
                            Exit For
                        End If
                    Next j
                End If
                'Close Metlite Workbook
                MetLite.Close (False)
                'If there is no CEI data found inform user
                If i = 49 Then
                    MsgBox "No CEI data found in Metlite file for " & Project.Text8, vbOKOnly, "Data Missing"
                End If
            ElseIf Not FS.FileExists(SearchFile) Then
                If MissingMetFile(0) = "" Then
                    MissingMetFile(UBound(MissingMetFile)) = Project.Text8
                ElseIf MissingMetFile(0) <> "" Then
                    ReDim Preserve MissingMetFile(UBound(MissingMetFile) + 1)
                    MissingMetFile(UBound(MissingMetFile)) = Project.Text8
                End If
            End If
        End If
    Next Project
    'delete empty columns
    MacroWB.Sheets(StatusMonth & "-" & StatusDay & "-" & StatusYear & " Metrics Analysis").Activate
    For j = ScheduleCount + 3 To 3 Step -1
        If Cells(3, j) = "" Then
            Columns(j).Delete
        End If
    Next j
    
    MsgBox "The following MetLite files could not be found: " & vbCrLf & vbCrLf & Join(MissingMetFile, vbCrLf)
    
    Call AppTrue
    
End Sub

'Generates the Contract Actions Analysis workbook
Sub GenerateAnalysis()

    'deactivate background tasks
    Call AppFalse
    
    Dim Project As CProject
    Dim CurrentCombined As New CCombinedProject
    Dim ContractAnalysisWB As Workbook
    Dim CombinedName As String
    
    'Set MS Project Object, Open Combined, Set combined object
    CombinedName = Dir(SchedulePath & "\" & "*Combined*")
    CurrentCombined.FileAddress = SchedulePath & "\" & CombinedName
    CurrentCombined.FileOpen
    CurrentCombined.GetSubProjects
    CurrentCombined.FileClose
    
    Dim j As Integer
    j = 3
    For Each Project In CurrentCombined.Projects
        Project.Text8 = MacroWB.Sheets(1).Cells(j, 1)
        j = j + 1
    Next Project
    
    Set ContractAnalysisWB = Workbooks.Add

    Dim CEIDataName As String
    Dim CEIDataWB As Workbook
    Dim CEISheet As Worksheet
    Dim ActionableName As String
    Dim BaselineSheet As Worksheet
    Dim BaselineDataWB As Workbook
    Dim NumberofStarts As Integer
    Dim ScheduleCount As Integer
    Dim ListRow As ListRow
    
    'variables for the table in excel
    Dim Identifier As String
    Dim Starts As Integer
    Dim Finishes As Integer
    Dim CriticalStarts As Integer
    Dim CriticalFinishes As Integer
    Dim NameLen As Integer
    Dim CEISheetName As String
    Dim BaselineSheetName As String
    
    'pull UID from comments to update risk charts
    Dim Range As Object
    Dim UID() As String
    Dim Field() As String
    Dim CommentCount As Integer
    Dim ListColumn As ListColumn
    Dim VarianceNeedRange As Range
    Dim CommentProject() As String
    Dim UniqueProject() As String
    Dim Unique As Boolean
    Dim oProject As Variant
    Dim UProject As Variant
    Dim index As Integer
    Dim CurrentProj As Project
    Dim FillData() As String
    Dim task As task
    Dim ContractAnalysisFilePath As String
    Dim i As Integer
    Dim TableColumn() As Variant
    Dim TableRow() As Variant
    Dim CurrentTask As Object
    Dim Columni As Variant
    Dim Rowi As Variant
    Dim CommentsProject() As String

    For Each Project In CurrentCombined.Projects
        
        CEISheetName = Project.Text8 & " CEI"
        NameLen = Len(Project.Text8 & " CEI")
        'concatenate Text 8 so that CEI is always in the Sheet name
        If NameLen > 31 Then
            CEISheetName = Left(Project.Text8, 27) & " CEI"
        End If
        
        CEIDataName = Dir(MetLitePath & "\" & Project.Text8 & "\" & StatusYear & "\" & StatusMonth & "\" & "*CEI DATA*.*")
        If CEIDataName <> "" Then
            'Add Sheet for CEI data
            Set CEISheet = ContractAnalysisWB.Sheets.Add
            ActiveSheet.Name = CEISheetName
            'copy CEI data from "Combination Results" sheet in MetLite
            Set CEIDataWB = Workbooks.Open(MetLitePath & "\" & Project.Text8 & "\" & StatusYear & "\" & StatusMonth & "\" & CEIDataName)
            CEIDataWB.Sheets("Combination Results").Activate
            ActiveSheet.Range(Cells(1, 1), Cells(FindLastRow, 25)).Copy
            'Paste data
            CEISheet.Paste
            'hide irrelevant columns
            For j = 1 To 25
                If j = 3 Or j = 4 Or j = 5 Or j = 6 Or j = 7 Or j = 8 Or j = 9 Or j = 12 Or j = 13 Or j = 14 Or _
                j = 15 Or j = 16 Or j = 17 Or j = 18 Or j = 19 Or j = 20 Or j = 21 Or j = 22 Or j = 23 Or j = 24 Then
                    CEISheet.Activate
                    CEISheet.Range(Cells(1, j), Cells(1, j)).Select
                    Selection.EntireColumn.Hidden = True
                End If
            Next j
            
            'autofit done on individual columns because an autofit on all was unhiding cells
            Columns(1).AutoFit
            Columns(2).AutoFit
            Columns(10).AutoFit
            Columns(11).AutoFit
            Columns(25).AutoFit
            
            CEIDataWB.Close
        'if the metlite file was not found
        ElseIf CEIDataName = "" Then
            Set CEISheet = ContractAnalysisWB.Sheets.Add
            
            ActiveSheet.Name = CEISheetName
            
        End If
        'create Baseline data sheet
        BaselineSheetName = Project.Text8 & " BL"
        NameLen = Len(Project.Text8 & " BL")
        'Concatenate Text 8 to always show BL
        If NameLen > 31 Then
            BaselineSheetName = Left(Project.Text8, 27) & " BL"
        End If
        
        'get data from Actionable Items workbook and pste into BL sheets
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ActionableName = Dir(MetLitePath & "\" & Project.Text8 & "\" & StatusYear & "\" & StatusMonth & "\" & "*Actionable*.*")
        If ActionableName <> "" Then
            
            BaselineSheetName = Project.Text8 & " BL"
            NameLen = Len(Project.Text8 & " BL")
            If NameLen > 31 Then
                BaselineSheetName = Left(Project.Text8, 27) & " BL"
            End If

            Set BaselineSheet = ContractAnalysisWB.Sheets.Add(Before:=ContractAnalysisWB.Sheets(Project.Text8 & " CEI"))
            
            ActiveSheet.Name = BaselineSheetName
            
            Set BaselineDataWB = Workbooks.Open(MetLitePath & "\" & Project.Text8 & "\" & StatusYear & "\" & StatusMonth & "\" & ActionableName)
            BaselineDataWB.Sheets("8-Delinquent Starts").Activate
            NumberofStarts = FindLastRow
            ActiveSheet.Range(Cells(1, 1), Cells(NumberofStarts, 25)).Copy
            
            BaselineSheet.Activate
            BaselineSheet.Cells(1, 1) = "Starts"
            BaselineSheet.Range(Cells(2, 1), Cells(2, 1)).Select
            Sheets(BaselineSheetName).Paste
            
            BaselineDataWB.Sheets("10-Delinquent Finishes").Activate
            ActiveSheet.Range(Cells(1, 1), Cells(FindLastRow, 25)).Copy
            
            BaselineSheet.Activate
            BaselineSheet.Cells(NumberofStarts + 3, 1) = "Finishes"
            BaselineSheet.Cells(NumberofStarts + 4, 1).Select
            Sheets(BaselineSheetName).Paste
            ActiveSheet.Range(Cells(1, 1), Cells(FindLastRow, FindLastCol)).Select
            Selection.Columns.AutoFit
            
            BaselineDataWB.Close (False)
            
        ElseIf ActionableName = "" Then
            
            Set BaselineSheet = ContractAnalysisWB.Sheets.Add(Before:=ContractAnalysisWB.Sheets(CEISheetName))
            ActiveSheet.Name = BaselineSheetName
            
        End If
        'end get data from Actionable Items workbook and pste into BL sheets'''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Next Project
    
    ContractAnalysisWB.Activate
    MacroWB.Sheets("Missed Tasks Breakdown").Copy Before:=ContractAnalysisWB.Sheets(1)
    ContractAnalysisWB.Sheets("Missed Tasks Breakdown").Activate

    MacroWB.Sheets("Risk").Copy Before:=ContractAnalysisWB.Sheets(1)
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'StartFill Variance to Need Data Table

    CommentCount = 0
    'select data body of table
    ContractAnalysisWB.Sheets("Risk").ListObjects("Table4").DataBodyRange.Select
    'select cells with comments within the databody of the table
    If Selection.SpecialCells(xlCellTypeComments).Count > 0 Then
        Selection.SpecialCells(xlCellTypeComments).Select
        'Cycle through the cells with comments
        For Each Rowi In Selection.Rows
            For Each Columni In Selection.Columns
                'set current cell as range
                Set Range = Cells(Rowi.Row, Columni.Column)
                    
                    CommentCount = CommentCount + 1
                    'capture uid in combined form, ie. text8-uid
                    ReDim Preserve UID(CommentCount - 1)
                    UID(CommentCount - 1) = Trim(Right(Range.Comment.Text, Len(Range.Comment.Text) - InStr(1, Range.Comment.Text, ":")))
                    'capture the project from the first part of the uid, ie text8
                    ReDim Preserve CommentsProject(CommentCount - 1)
                    CommentsProject(CommentCount - 1) = Left(UID(CommentCount - 1), Len(UID(CommentCount - 1)) - (Len(UID(CommentCount - 1)) - InStr(1, UID(CommentCount - 1), "-") + 1))
                    'capture the column to which to return the value
                    ReDim Preserve TableColumn(CommentCount - 1)
                    TableColumn(CommentCount - 1) = Columni.Column
                    'capture the row ot which to return the value
                    ReDim Preserve TableRow(CommentCount - 1)
                    TableRow(CommentCount - 1) = Rowi.Row
                    
            Next Columni
        Next Rowi
        
        'find number of unique projects with comments
        ReDim UniqueProject(0)
        UniqueProject(0) = Trim(CommentsProject(0))
        'set the size of the Table data array (used to return the values to the table)
        ReDim TableData(CommentCount - 1, 2) As Variant
        'cycle through the comments to Comments project array to find unique entries
        For i = 0 To UBound(CommentsProject)
            Unique = True
            For j = 0 To UBound(UniqueProject)
                If CommentsProject(i) = UniqueProject(j) Then
                    Unique = False
                End If
            Next j
            If Unique = True Then
                'add unique entries to the Unique Project array
                ReDim Preserve UniqueProject(UBound(UniqueProject) + 1)
                UniqueProject(UBound(UniqueProject)) = Trim(CommentsProject(i))
            End If
        Next i
        
        'open the file for each unique project and get the task field to be entered into the table
        For Each UProject In UniqueProject()
            For Each Project In CurrentCombined.Projects
                If CStr(UProject) = Project.Text8 Then
                    Project.FileOpen
                    i = 0
                    For Each oProject In CommentsProject()
                        If oProject = Project.Text8 Then
                            Set CurrentTask = appProj.ActiveProject.Tasks.UniqueID(CInt(Right(UID(i), Len(UID(i)) - InStr(1, UID(i), "-"))))
                            TableData(i, 0) = TableRow(i)
                            TableData(i, 1) = TableColumn(i)
                            TableData(i, 2) = CurrentTask.Finish
                            i = i + 1
                        End If
                    Next oProject
                    Project.FileClose
                End If
            Next Project
        Next UProject
        'enter the information from the tabledata array into the table
        ContractAnalysisWB.Activate
        Sheets("Risk").Activate
        For i = 0 To UBound(TableData, 1)
            Cells(TableData(i, 0), TableData(i, 1)) = TableData(i, 2)
        Next i
        Application.Calculate
    End If
    
    
    'End Fill Variance to Need Data Table
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Start Fill in Finish Variance Table
    
    'erase arrays to reuse them
    Erase UID
    Erase CommentsProject
    Erase TableColumn
    Erase TableRow
    Erase UniqueProject
    
    CommentCount = 0
    'select data body of table
    ContractAnalysisWB.Sheets("Risk").ListObjects("Table5").DataBodyRange.Select
    'select cells with comments within the databody of the table
    If Selection.SpecialCells(xlCellTypeComments).Count > 0 Then
        Selection.SpecialCells(xlCellTypeComments).Select
        
        'Cycle through the cells with comments
        For Each Rowi In Selection.Rows
            For Each Columni In Selection.Columns
                'set current cell as range
                Set Range = Cells(Rowi.Row, Columni.Column)
                    
                    CommentCount = CommentCount + 1
                    'capture uid in combined form, ie. text8-uid
                    ReDim Preserve UID(CommentCount - 1)
                    UID(CommentCount - 1) = Trim(Right(Range.Comment.Text, Len(Range.Comment.Text) - InStr(1, Range.Comment.Text, ":")))
                    'capture the project from the first part of the uid, ie text8
                    ReDim Preserve CommentsProject(CommentCount - 1)
                    CommentsProject(CommentCount - 1) = Left(UID(CommentCount - 1), Len(UID(CommentCount - 1)) - (Len(UID(CommentCount - 1)) - InStr(1, UID(CommentCount - 1), "-") + 1))
                    'capture the column to which to return the value
                    ReDim Preserve TableColumn(CommentCount - 1)
                    TableColumn(CommentCount - 1) = Columni.Column
                    'capture the row ot which to return the value
                    ReDim Preserve TableRow(CommentCount - 1)
                    TableRow(CommentCount - 1) = Rowi.Row
                    
            Next Columni
        Next Rowi
        
        'find number of unique projects with comments
        ReDim UniqueProject(0)
        UniqueProject(0) = Trim(CommentsProject(0))
        'set the size of the Table data array (used to return the values to the table)
        ReDim TableData(CommentCount - 1, 2) As Variant
        'cycle through the comments to Comments project array to find unique entries
        For i = 0 To UBound(CommentsProject)
            Unique = True
            For j = 0 To UBound(UniqueProject)
                If CommentsProject(i) = UniqueProject(j) Then
                    Unique = False
                End If
            Next j
            If Unique = True Then
                'add unique entries to the Unique Project array
                ReDim Preserve UniqueProject(UBound(UniqueProject) + 1)
                UniqueProject(UBound(UniqueProject)) = Trim(CommentsProject(i))
            End If
        Next i
        
        'open the file for each unique project and get the task field to be entered into the table
        For Each UProject In UniqueProject()
            For Each Project In CurrentCombined.Projects
                If CStr(UProject) = Project.Text8 Then
                    Project.FileOpen
                    i = 0
                    For Each oProject In CommentsProject()
                        If oProject = Project.Text8 Then
                            Set CurrentTask = appProj.ActiveProject.Tasks.UniqueID(CInt(Right(UID(i), Len(UID(i)) - InStr(1, UID(i), "-"))))
                            TableData(i, 0) = TableRow(i)
                            TableData(i, 1) = TableColumn(i)
                            'check the column from the table and get the correct field accordingly
                            If TableColumn(i) = 4 Then
                                TableData(i, 2) = CurrentTask.BaselineStart
                            ElseIf TableColumn(i) = 5 Then
                                TableData(i, 2) = CurrentTask.BaselineFinish
                            End If
                            
                            i = i + 1
                        End If
                    Next oProject
                    Project.FileClose
                End If
            Next Project
        Next UProject
        
        'enter the information from the tabledata array into the table
        ContractAnalysisWB.Activate
        Sheets("Risk").Activate
        For i = 0 To UBound(TableData, 1)
            Cells(TableData(i, 0), TableData(i, 1)) = TableData(i, 2)
        Next i
        Application.Calculate
    
    End If
    
    
    'End Fill in Finish Variance Table
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Start Fill in Missed Tasks Baseline Chart
    Erase TableData
    ReDim TableData(CurrentCombined.Projects.Count, 5)
    Dim Sheet As Worksheet
    Dim k As Integer
    Dim l As Integer
    Dim m As Integer
    Dim CurrentProjectIndex As Integer
    
    'TableData Fields: row, first column, starts, finishes, critical starts, critical finishes
    
    ContractAnalysisWB.Activate
    Sheets("Missed Tasks Breakdown").Activate
    'add a row to the table for each project - the one that is already there
    For j = 1 To CurrentCombined.Projects.Count - 1
        ContractAnalysisWB.Sheets("Missed Tasks Breakdown").Range("Table6").ListObject.ListRows.Add (2)
    Next j
    
    ContractAnalysisWB.Sheets("Missed Tasks Breakdown").ListObjects("Table6").DataBodyRange.Select
    
    i = 1
    'fill in te contract names in the table
    For Each Project In CurrentCombined.Projects
        Cells(Selection.Rows(i).Row, Selection.Columns(1).Column) = Project.Text8
        TableData(i - 1, 0) = Selection.Rows(i).Row
        TableData(i - 1, 1) = Selection.Columns(1).Column + 1
        i = i + 1
    Next Project
    
    CurrentProjectIndex = -1
    
    For Each Project In CurrentCombined.Projects
        CurrentProjectIndex = CurrentProjectIndex + 1
        For Each Sheet In ContractAnalysisWB.Sheets
            'find the baseline sheet for that project
            If InStr(1, Sheet.Name, Project.Text8) > 0 And InStr(1, Sheet.Name, "BL") > 0 Then
                Sheet.Activate
                'if there is no data then exit for
                If Cells(1, 1) = "" Then
                    TableData(CurrentProjectIndex, 2) = ""
                    TableData(CurrentProjectIndex, 3) = ""
                    TableData(CurrentProjectIndex, 4) = ""
                    TableData(CurrentProjectIndex, 5) = ""
                    Exit For
                End If
                'count the starts, finishes, critical starts, critical finishes
                j = 3
                While Cells(j, 1) <> "Finishes"
                    j = j + 1
                Wend
                
                TableData(CurrentProjectIndex, 2) = j - 4
                
                CriticalStarts = 0
                For l = 3 To j - 2
                    If Cells(l, 6) = 0 Then
                        CriticalStarts = CriticalStarts + 1
                    End If
                Next l
                
                TableData(CurrentProjectIndex, 4) = CriticalStarts
                
                k = j
                While Cells(k, 1) <> ""
                    k = k + 1
                Wend
                
                TableData(CurrentProjectIndex, 3) = k - j - 2
                
                CriticalFinishes = 0
                For m = j + 2 To k - 1
                    If Cells(m, 6) = 0 Then
                        CriticalFinishes = CriticalFinishes + 1
                    End If
                Next m
                
                TableData(CurrentProjectIndex, 5) = CriticalFinishes
                
            End If
        Next Sheet
    Next Project
    
    'fill in table from TableData array
    Sheets("Missed Tasks Breakdown").Activate
    For i = 0 To CurrentCombined.Projects.Count - 1
        For j = 0 To 3
            Cells(TableData(i, 0), TableData(i, 1) + j) = TableData(i, j + 2)
        Next j
    Next i
    
    'End Fill Missed Baseline Tasks Chart
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Start Fill CEI Chart
    
    'erase tabledata to reuse
    Erase TableData
    ReDim TableData(CurrentCombined.Projects.Count, 2)
    Dim Row As Object
    
    
    'TableData: Tasks that did finish, tasks that did not finish, Text8
    'Add rows to table for ach project
    For j = 1 To CurrentCombined.Projects.Count - 1
        ContractAnalysisWB.Sheets("Missed Tasks Breakdown").Range("Table37").ListObject.ListRows.Add (2)
    Next j
    
    ContractAnalysisWB.Sheets("Missed Tasks Breakdown").ListObjects("Table37").DataBodyRange.Select
    
    i = 1
    'fill in project names
    For Each Project In CurrentCombined.Projects
        Cells(Selection.Rows(i).Row, Selection.Columns(1).Column) = Project.Text8
        i = i + 1
    Next Project
    
    MacroWB.Sheets(StatusMonth & "-" & StatusDay & "-" & StatusYear & " Metrics Analysis").Activate
    
    j = 3
    While Cells(89, j) <> "Total of individual"
        TableData(j - 3, 0) = Cells(90, j)
        TableData(j - 3, 1) = Cells(91, j)
        TableData(j - 3, 2) = Trim(Left(Cells(3, j), Len(Cells(3, j)) - 11))
        j = j + 1
    Wend
    
    ContractAnalysisWB.Sheets("Missed Tasks Breakdown").Activate
    ContractAnalysisWB.Sheets("Missed Tasks Breakdown").ListObjects("Table37").DataBodyRange.Select
    
    j = 0
    For Each Row In Selection.Rows
        If Cells(Row.Row, Selection.Columns(1).Column) = TableData(j, 2) Then
        
            Cells(Row.Row, Selection.Columns(2).Column) = TableData(j, 0)
            Cells(Row.Row, Selection.Columns(3).Column) = TableData(j, 1)
            j = j + 1
        End If
    Next Row
    
    'End Fill CEI Chart
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Start Fill in Summary Table
    'this range should probably be truned into a table, it would be much easier to manipulate
    
    Erase TableData
    ReDim TableData(CurrentCombined.Projects.Count, 0)
    Dim Cell As Cell
    Dim Count As Integer
    
    MacroWB.Sheets(StatusMonth & "-" & StatusDay & "-" & StatusYear & " Metrics Analysis").Activate
    
    j = 3
    While Cells(3, j) <> "Total of individual"
        TableData(j - 3, 0) = Cells(44, j)
        j = j + 1
    Wend
    
    ContractAnalysisWB.Sheets("Missed Tasks Breakdown").Activate
    
    For j = 1 To UBound(TableData)
        Cells(j + 3, 13) = TableData(j - 1, 0)
    Next j
    
    j = 4
    Count = CInt(CurrentCombined.Projects.Count)
    
    ContractAnalysisWB.Sheets("Missed Tasks Breakdown").Range(Cells(4, 12), Cells(Count + 3, 12)) = ""

    For Each Project In CurrentCombined.Projects
        Cells(j, 12) = Cells(j, 8)
        j = j + 1
    Next Project
    
    ContractAnalysisWB.Sheets("Risk").Activate
    ContractAnalysisWB.Sheets("Risk").ListObjects("Table5").DataBodyRange.Select
    
    For j = 1 To UBound(TableData)
        TableData(j - 1, 0) = Cells(j + 25, Selection.Columns.Count - 1) ' coulmns.count - 1 gives calendar day variance because the table starts in column2
    Next j
    
    ContractAnalysisWB.Sheets("Missed Tasks Breakdown").Activate
    
    For j = 1 To UBound(TableData)
        Cells(j + 3, 16) = TableData(j - 1, 0)
    Next j
    
    MacroWB.Sheets(StatusMonth & "-" & StatusDay & "-" & StatusYear & " Metrics Analysis").Activate
    
    j = 3
    While Cells(89, j) <> "Total of individual"
        TableData(j - 3, 0) = Cells(93, j)
        j = j + 1
    Wend
    
    ContractAnalysisWB.Sheets("Missed Tasks Breakdown").Activate
    
    For j = 1 To UBound(TableData)
        Cells(j + 3, 17) = TableData(j - 1, 0)
    Next j
    
    'End Fill in Summary Table
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    Call AppTrue
    
         
    If Not FS.FolderExists(TopFolderPath & "\" & "Monthly Products") Then
        MkDir (TopFolderPath & "\" & "Monthly Products")
    End If
    
    ScheduleCount = 30
    
    While MacroWB.Sheets(1).Cells(ScheduleCount, 1) = ""
    ScheduleCount = ScheduleCount - 1
    Wend
    
    For i = 2 To ScheduleCount
        If Not FS.FolderExists(TopFolderPath & "\" & "Monthly Products" & "\" & MacroWB.Sheets(1).Cells(i, 1)) Then
            MkDir (TopFolderPath & "\" & "Monthly Products" & "\" & MacroWB.Sheets(1).Cells(i, 1))
        End If
        If Not FS.FolderExists(TopFolderPath & "\" & "Monthly Products" & "\" & MacroWB.Sheets(1).Cells(i, 1) & "\" & StatusYear) Then
            MkDir (TopFolderPath & "\" & "Monthly Products" & "\" & MacroWB.Sheets(1).Cells(i, 1) & "\" & StatusYear)
        End If
        If Not FS.FolderExists(TopFolderPath & "\" & "Monthly Products" & "\" & MacroWB.Sheets(1).Cells(i, 1) & "\" & StatusYear & "\" & StatusMonth) Then
            MkDir (TopFolderPath & "\" & "Monthly Products" & "\" & MacroWB.Sheets(1).Cells(i, 1) & "\" & StatusYear & "\" & StatusMonth)
        End If
    Next i
    
    ContractAnalysisFilePath = TopFolderPath & "\" & "Monthly Products" & "\" & "Combined" & "\" & StatusYear & "\" & StatusMonth & "\" & "Contract Action Analysis " & StatusMonth & "-" & StatusDay & "-" & StatusYear
    
    If FS.FileExists(ContractAnalysisFilePath & ".xlsx") Then
        Kill (ContractAnalysisFilePath & "*")
    End If
    
    ContractAnalysisWB.SaveAs (ContractAnalysisFilePath)
    ContractAnalysisWB.Close (True)
    appProj.FileExit (pjDoNotSave)
     
    MacroWB.Sheets("Input").Activate
    MsgBox "Contract Actions Analysis Workbook Generated"
    
    Call AppTrue
            
End Sub
'does not work at all
Sub PopulatePowerPoint()
    Dim MonthlyCombined As CCombinedProject
    MonthlyCombined.GetSubProjects (True)
    Dim Project As CProject
    Dim FileAddress As String
    Dim Filename As String

    For Each Project In MonthlyCombined.Projects
        If StatusMonth > 1 Then
            FileAddress = TopFolderPath & "\Monthly Products\" & Project.Text8 & "\" & StatusYear & "\" & StatusMonth - 1 & "\"
        ElseIf StatusMonth = 1 Then
            FileAddress = TopFolderPath & "\Monthly Products\" & Project.Text8 & "\" & StatusYear - 1 & "\" & 12 & "\"
        End If

'        Filename = Dir(FileAddress & ")
        FileAddress = FileAddress & Filename

'        ShellExecute 0, vbNullString, FileAddress, 0&, 0&, 1
'
'        Project.RawData =
'
    
    
    
    
    
End Sub
'not sure if this is ever actually used
Function SaveAs(initialFilename As String, StatusDate)
  SaveAs = False
  With Application.FileDialog(msoFileDialogSaveAs)
    .AllowMultiSelect = False
    .ButtonName = "&Save As"
    .initialFilename = initialFilename
    .title = "File Save As ("
    .Show
    SaveAs = .SelectedItems(1)
  End With
End Function
'not sure if this is used either
Function FindNumberSched()
    If WorksheetFunction.CountA(Cells) > 0 Then
        
        FindNumberSched = ActiveWorkbook.Sheets(1).Cells.Find(What:="Total of individual", After:=[A1], _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column - 4
    End If
End Function

' Find last row on sheet
Function FindLastRow() As Integer
    
    If WorksheetFunction.CountA(Cells) > 0 Then
        FindLastRow = Cells.Find(What:="*", After:=[A1], _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    End If
         
End Function

' Find last Column on sheet
Function FindLastCol() As Integer
    
    If WorksheetFunction.CountA(Cells) > 0 Then
        FindLastCol = Cells.Find(What:="*", After:=[A1], _
        SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    End If
    
End Function
' this might be used somewhere
Function GetFileName(Path) As String
    Dim Last As Integer
    Dim First As Integer
    Dim Filename As String
    
    Last = InStrRev(Path, ".")
    First = InStrRev(Path, "\") - 1
    Filename = Left(Path, Last - 1)
    Filename = Right(Path, Len(Filename) - First - 1)
    
    GetFileName = Filename
End Function
Function RightOf(SourceString, Symbol)
    RightOf = Right(SourceString, Len(SourceString) - InStrRev(SourceString, Symbol))
End Function
Function LeftOf(SourceString, Symbol)
    LeftOf = Left(SourceString, Len(SourceString) - (Len(SourceString) - (InStr(SourceString, Symbol) - 1)))
End Function

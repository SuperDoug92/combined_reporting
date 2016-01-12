Attribute VB_Name = "PublicData"
Public MacroWB As Workbook
Public appProj As MSProject.Application
Public CombinedProj As MSProject.Project
Public CombinedRawDataWB As Workbook
Public StatusDate As String
Public StatusMonth As String
Public StatusDay As String
Public StatusYear As String
Public TopFolderPath As String
Public MetAnalWB As Workbook
Public CAAnalWB As Workbook
Public CombProjAd As String
Public MetAd As String
Public Schedules()
Public MetLitePath As String
Public FS As Object
Public SchedulePath As String
Public EOMSchedulePath As String

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

Sub SetUp()
    
    Call AppFalse
    Set appProj = CreateObject("MSProject.Application")
    Set FS = CreateObject("Scripting.FileSystemObject")
    'Set the Macro Workbook Object
    Set MacroWB = ActiveWorkbook
    'Define the Status Date and Metlite version name variables from cells in the macro workbook
    If Len(ActiveWorkbook.Sheets(1).Cells(2, 3)) < 10 Then
        StatusDate = 0 & ActiveWorkbook.Sheets(1).Cells(2, 3)
    Else
        StatusDate = ActiveWorkbook.Sheets(1).Cells(2, 3)
    End If
    StatusMonth = Left(StatusDate, 2)
    StatusDay = Mid(StatusDate, 4, 2)
    StatusYear = Right(StatusDate, 4)
    TopFolderPath = ActiveWorkbook.Sheets(1).Cells(3, 4)
    MetLitePath = TopFolderPath & "\MetLite"
    SchedulePath = TopFolderPath & "\# SCHEDULES"
    EOMSchedulePath = SchedulePath & "\" & StatusYear & "\" & StatusMonth

    Call AppTrue
    
End Sub

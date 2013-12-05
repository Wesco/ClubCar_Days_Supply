Attribute VB_Name = "Program"
Option Explicit
Public Const VersionNumber As String = "1.0.0"

Sub Main()
    ImportGaps
    ImportMaster
    ImportForecast
    CreateDSReport
    ExportReport
    Clean
    
    MsgBox "Complete!", Title:="Macro"
End Sub

Sub Clean()
    Dim PrevDispAlert As Boolean
    Dim PrevScrnUpdat As Boolean
    Dim PrevWkbk As Workbook
    Dim s As Worksheet
    
    Set PrevWkbk = ActiveWorkbook
    PrevDispAlert = Application.DisplayAlerts
    PrevScrnUpdat = Application.ScreenUpdating
    ThisWorkbook.Activate
    
    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            Cells.Delete
            Range("A1").Select
        End If
    Next
    
    Sheets("Macro").Select
    Range("C7").Select
    
    PrevWkbk.Activate
    Application.DisplayAlerts = PrevDispAlert
    Application.ScreenUpdating = PrevScrnUpdat
End Sub

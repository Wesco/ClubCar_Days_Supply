Attribute VB_Name = "Export"
Option Explicit

Sub ExportReport()
    Dim PrevDispAlert As Boolean
    Dim Path As String
    Dim File As String
    Dim Msg As String

    PrevDispAlert = Application.DisplayAlerts
    Path = "\\br3615gaps\gaps\Club Car\Days Supply\" & Format(Date, "yyyy") & "\"
    File = "Days Supply " & Format(Date, "yyyy-mm-dd") & ".xlsx"

    If Not FolderExists(Path) Then
        RecMkDir Path
    End If

    Sheets("Combined").Copy
    ActiveSheet.Name = "Days Supply"

    Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Path & File, xlOpenXMLWorkbook
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlert

    Msg = "The days supply report is attached. A copy can also be found on the network <a href=""" & Path & File & """>here</a>."
    Email SendTo:=Environ("username") & "@wesco.com", Body:=Msg, Attachment:=Path & File
End Sub

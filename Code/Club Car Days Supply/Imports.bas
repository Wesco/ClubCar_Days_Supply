Attribute VB_Name = "Imports"
Option Explicit

Sub ImportMaster()
    Dim Path As String
    Dim File As String

    Path = "\\br3615gaps\gaps\Club Car\Master\"
    File = "Club Car Master " & Format(Date, "yyyy") & ".xlsx"

    If FileExists(Path & File) Then
        ImportFile Path, File, ThisWorkbook.Sheets("Master").Range("A1")
    Else
        Err.Raise Errors.FILE_NOT_FOUND, "ImportMaster", "The Club Car master was not imported."
    End If

End Sub

Sub ImportForecast()
    Dim Imported As Boolean
    Dim TotalRows As Long
    Dim Path As String
    Dim FcstA As String
    Dim FcstP As String
    Dim i As Integer
    Dim dt As Date

    Sheets("Combined").Select

    For i = 0 To 30
        dt = Date - i
        Path = "\\br3615gaps\gaps\Club Car\Forecast\" & Format(Date, "yyyy") & "\"
        FcstA = "Warehouse A forecast " & Format(dt, "mm-dd-yy")
        FcstP = "Warehouse P forecast " & Format(dt, "mm-dd-yy")

        If FileExists(Path & FcstA & ".xlsx") And FileExists(Path & FcstP & ".xlsx") Then
            If i > 0 Then
                If MsgBox("A forecast from " & Format(dt, "mmm dd, yyyy") & " was found." & vbCrLf & _
                          "Would you like to import this file?", vbYesNo, "Import old forecast?") = vbYes Then
                    ImportFile Path, FcstA & ".xlsx", ThisWorkbook.Sheets("Combined").Range("A1")
                    TotalRows = ActiveSheet.UsedRange.Rows.Count + 1
                    ImportFile Path, FcstP & ".xlsx", ThisWorkbook.Sheets("Combined").Range("A" & TotalRows)
                    Rows(TotalRows).Delete
                    Imported = True
                End If
                Exit For
            Else
                ImportFile Path, FcstA & ".xlsx", ThisWorkbook.Sheets("Combined").Range("A1")
                TotalRows = ActiveSheet.UsedRange.Rows.Count + 1
                ImportFile Path, FcstP & ".xlsx", ThisWorkbook.Sheets("Combined").Range("A" & TotalRows)
                Rows(TotalRows).Delete
                Imported = True
                Exit For
            End If
        End If
    Next

    If Imported = False Then
        Err.Raise Errors.FILE_NOT_FOUND, "ImportForecast", "A Club Car forecast was not imported."
    End If
End Sub

Private Sub ImportFile(Path As String, FileName As String, Destination As Range)
    Dim PrevDispAlert As Boolean

    PrevDispAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False

    Workbooks.Open Path & FileName
    ActiveSheet.UsedRange.Copy Destination:=Destination
    ActiveWorkbook.Close

    Application.DisplayAlerts = PrevDispAlert
End Sub

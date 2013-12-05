Attribute VB_Name = "CreateReport"
Option Explicit

Sub CreateDSReport()
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim ColHeaders As Variant

    Sheets("Combined").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    'Remove Total column
    Columns(TotalCols).Delete
    TotalCols = ActiveSheet.UsedRange.Columns.Count

    'Sort Item Number A-Z
    With ActiveWorkbook.Worksheets("Combined").Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("A1"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .SetRange Range("A2:N" & TotalRows)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'Subtotal by Item Number
    ActiveSheet.UsedRange.Subtotal GroupBy:=1, _
                                   Function:=xlSum, _
                                   TotalList:=Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14), _
                                   Replace:=True, _
                                   PageBreaks:=False, _
                                   SummaryBelowData:=True

    'Store subtotals as values
    Cells.Copy
    Range("A1").PasteSpecial Paste:=xlPasteValues, _
                             Operation:=xlNone, _
                             SkipBlanks:=False, _
                             Transpose:=False
    Application.CutCopyMode = False

    'Remove subtotal formatting
    ActiveSheet.UsedRange.RemoveSubtotal
    Range("A1").Select

    'Remove raw data
    ActiveSheet.UsedRange.AutoFilter Field:=1, Criteria1:="<>*Total*"
    ColHeaders = Range("A1:N1")
    Cells.Delete
    Rows(1).Insert
    Range("A1:N1").Value = ColHeaders
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Remove Gand Total
    Rows(TotalRows).Delete

    'Get TotalRows
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Remove postfix
    Range("A1:A" & TotalRows).Replace " Total", ""

    'Remove description column
    Columns("B:B").Delete

    'Insert SIM column
    Columns("B:B").Insert

    'Lookup SIM
    Range("B1").Value = "SIM"
    Range("B2:B" & TotalRows).Formula = "=IFERROR(VLOOKUP(A2,Master!A:B,2,FALSE),"""")"
    Range("B2:B" & TotalRows).NumberFormat = "@"
    Range("B2:B" & TotalRows).Value = Range("B2:B" & TotalRows).Value

    'Add Qty Needed Per Day
    Range("N1").Value = "Days On Hand"
    Range("N2:N" & TotalRows).Formula = "=IFERROR(VLOOKUP(B2,Gaps!A:G,7,FALSE)/(SUM(C2:M2)/236),0)"
    Range("N2:N" & TotalRows).Value = Range("N2:N" & TotalRows).Value
    
    'Remove monthly quantities
    Columns("C:M").Delete
    
    'Fix formatting
    ActiveSheet.UsedRange.Font.Bold = False
    Range("A1:C1").Font.Bold = True
End Sub

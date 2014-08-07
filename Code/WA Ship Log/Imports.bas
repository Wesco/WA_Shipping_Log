Attribute VB_Name = "Imports"
Option Explicit

Sub PORImport()
    On Error GoTo Import_Error

    'Prompt user to import POR report, delete it after it is imported
    UserImportFile Sheets("POR").Range("A1"), False

    'Clean up the report
    Sheets("POR").Select
    Rows(1).Delete
    Columns(Columns.Count).End(xlToLeft).Delete

    On Error GoTo 0
    Exit Sub

Import_Error:
    Err.Raise Errors.USER_INTERRUPT, "ImportPOR", "POR Import Canceled"
    Exit Sub
End Sub

Sub MasterImport()
    Dim PrevDispAlerts As Boolean
    Dim Path As String
    Dim TotalRows As Long
    Dim i As Long

    PrevDispAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    Path = "\\br3615gaps\gaps\Club Car\Master\Club Car Master " & Format(Date, "yyyy") & ".xlsx"

    'Import Club Cars master list
    Workbooks.Open Path
    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Master").Range("A1")
    ActiveWorkbook.Close

    Sheets("Master").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Convert part numbers to strings
    For i = 1 To TotalRows
        Cells(i, 1).Formula = "=""" & Cells(i, 1).Value & """"
    Next
    Range("A1:A" & TotalRows).NumberFormat = "@"

    'Copy the part numbers to the right of SIMs
    'so that a vlookup from SIM to Part is possible
    Columns(3).Insert
    Range("A1:A" & TotalRows).Copy Destination:=Range("C1")

    Application.DisplayAlerts = PrevDispAlerts
End Sub

Sub KitBOMImport()
    Dim ColHeaders As Variant
    Dim PrevDispAlerts As Boolean
    Dim Path As String
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim i As Long

    Path = "\\br3615gaps\gaps\Club Car\Master\Kit BOM " & Format(Date, "yyyy") & ".xlsx"

    'Import Kit BOM
    Workbooks.Open Path
    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Kit BOM").Range("A1")
    ActiveWorkbook.Close

    Sheets("Kit BOM").Select
    TotalCols = ActiveSheet.UsedRange.Columns.Count
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Copy column headers to the row directly above data
    Range(Cells(4, 1), Cells(4, TotalCols)).Value = Range(Cells(2, 1), Cells(2, TotalCols)).Value

    'Remove superfluous rows above column headers
    Rows("1:3").Delete

    'Remove quotes surrounding SIM numbers
    'Kit SIMs may contain additional spaces, by removing the surrounding quotes
    'these spaces are removed and SIMs get stored as numbers. This should not cause
    'any leading 0's to be lost because all kits begin with 9
    Range(Cells(2, 3), Cells(TotalRows, 3)).Replace "'", "", SearchOrder:=xlByRows

    'Store SIMs as strings
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Columns(3).Insert
    Range("C1").Value = "SIM"
    With Range(Cells(2, 3), Cells(TotalRows, 3))
        .Formula = "=""=""&""""""""&D2&"""""""""
        .Value = .Value
    End With
    Columns(4).Delete

    'Store component SIMs as strings
    Columns(6).Insert
    Range("F1").Value = "Comp SIM"
    Range("F2:F" & TotalRows).Formula = "=""=""""""&SUBSTITUTE(TRIM(G2),""'"","""")&"""""""""
    With Range(Cells(2, 6), Cells(TotalRows, 6))
        .Value = .Value
        .Replace " ", "", SearchOrder:=xlByRows
    End With
    Columns(7).Delete

    'Remove all record types except I & J
    TotalCols = ActiveSheet.UsedRange.Columns.Count
    ColHeaders = Range(Cells(1, 1), Cells(1, TotalCols))
    ActiveSheet.UsedRange.AutoFilter Field:=5, Criteria1:="<>J", Operator:=xlAnd, Criteria2:="<>I"
    Cells.Delete

    'Add column headers back to the report
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, TotalCols)) = ColHeaders
End Sub

Sub GapsImport()
    Dim TotalRows As Long

    'Import GAPs
    ImportGaps Sheets("Gaps").Range("A1"), True

    Sheets("Gaps").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Remove all columns except SIM & SIM Description
    Columns("G:CV").Delete
    Columns("B:E").Delete
End Sub

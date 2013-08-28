Attribute VB_Name = "Imports"
Option Explicit

Sub PORImport()
    On Error GoTo Import_Error
    'Change False to True before publishing macro
    UserImportFile Sheets("POR").Range("A1"), False

    Sheets("POR").Select
    Rows(1).Delete
    Columns(Columns.Count).End(xlToLeft).Delete
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
    Path = "\\br3615gaps\gaps\Club Car\Master\Club Car Master 2013.xlsx"

    If FileExists(Path) Then
        Workbooks.Open Path
        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Master").Range("A1")
        ActiveWorkbook.Close

        Sheets("Master").Select
        TotalRows = ActiveSheet.UsedRange.Rows.Count

        For i = 1 To TotalRows
            Cells(i, 1).Formula = "=""" & Cells(i, 1).Value & """"
        Next

        Range("A1:A" & TotalRows).NumberFormat = "@"
        Columns(3).Insert
        Range("A1:A" & TotalRows).Copy Destination:=Range("C1")

        Application.DisplayAlerts = PrevDispAlerts
    Else
        Application.DisplayAlerts = PrevDispAlerts
        Err.Raise Errors.FILE_NOT_FOUND, "ImportMaster", "File not found"
    End If
End Sub

Sub KitBOMImport()
    Dim ColHeaders As Variant
    Dim PrevDispAlerts As Boolean
    Dim Path As String
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim i As Long

    Path = "\\br3615gaps\gaps\Club Car\Master\Kit BOM 2013.xlsx"

    Workbooks.Open Path
    ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Kit BOM").Range("A1")
    ActiveWorkbook.Close

    Sheets("Kit BOM").Select
    TotalCols = ActiveSheet.UsedRange.Columns.Count
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    Range(Cells(4, 1), Cells(4, TotalCols)).Value = Range(Cells(2, 1), Cells(2, TotalCols)).Value
    Rows("1:3").Delete
    Range(Cells(2, 3), Cells(TotalRows, 3)).Replace "'", "", SearchOrder:=xlByRows
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    Columns(3).Insert
    Range("C1").Value = "SIM"
    With Range(Cells(2, 3), Cells(TotalRows, 3))
        .Formula = "=""=""&""""""""&D2&"""""""""
        .Value = .Value
    End With
    Columns(4).Delete

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
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, TotalCols)) = ColHeaders
End Sub

Sub GapsImport()
    Dim TotalRows As Long
    Dim TotalCols As Integer
    
    ImportGaps

    Sheets("Gaps").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count
    
    With Range("A2:A" & TotalRows)
        .ClearContents
        .Formula = "=""=""""""&C2&D2&"""""""""
        .Value = .Value
    End With
    
    Columns("G:CV").Delete
    Columns("B:E").Delete
End Sub






























Attribute VB_Name = "Program"
Option Explicit
Public Const VersionNumber As String = "2.0.1"
Public POCount As Integer

'---------------------------------------------------------------------------------------
' Proc : CreateShipment
' Date : 12/27/2013
' Desc : Gets all of the shipment data
'---------------------------------------------------------------------------------------
Sub CreateShipment()
    Dim NumOfPOs As Integer
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim ColHeaders() As Variant
    Dim POList() As Long
    Dim RowCount As Long
    Dim PO As String
    Dim i As Long
    Dim j As Long

    Application.ScreenUpdating = False
    Clean
    
    'Prompt user to import Purchase Order Report
    On Error GoTo Import_Error
    PORImport

    'Import Master
    MasterImport

    'Import Gaps
    GapsImport

    'Import Kit BOM
    KitBOMImport

    'Prompt user for the number of POs on the shipment
    frmNumPOs.Show
    NumOfPOs = frmNumPOs.NumPOs
    Unload frmNumPOs
    If NumOfPOs = 0 Then
        Err.Raise Errors.USER_INTERRUPT
    End If
    On Error GoTo 0

    'Dimension POList array
    ReDim POList(1 To NumOfPOs) As Long

    'Prompt user for PO numbers
    'Copy the PO that was entered onto Ship Log
    Sheets("POR").Select
    ReDim POList(1 To NumOfPOs) As Long
    For i = 1 To NumOfPOs
        POCount = i
        frmGetPO.Show
        PO = frmGetPO.PO
        Unload frmGetPO

        On Error GoTo Import_Error
        If PO = 0 Then
            Err.Raise Errors.USER_INTERRUPT
        End If
        On Error GoTo 0

        POList(i) = PO
        ActiveSheet.UsedRange.AutoFilter 3, "=" & PO, xlAnd
        RowCount = RowCount + 1
        ActiveSheet.UsedRange.Copy Destination:=Sheets("Ship Log").Range("A" & RowCount)
        RowCount = Sheets("Ship Log").UsedRange.Rows.Count
    Next

    'Remove columns that are not needed
    Sheets("Ship Log").Select
    For i = ActiveSheet.UsedRange.Columns.Count To 1 Step -1
        If Cells(1, i).Value <> "PO NUMBER" And _
           Cells(1, i).Value <> "DESCRIPTION" And _
           Cells(1, i).Value <> "QTY ORD" And _
           Cells(1, i).Value <> "ORDER" And _
           Cells(1, i).Value <> "LINE" And _
           Cells(1, i).Value <> "PRICE" Then
            Columns(i).Delete
        End If
    Next

    'Remove duplicate column header rows
    TotalCols = ActiveSheet.UsedRange.Columns.Count
    ColHeaders = Range(Cells(1, 1), Cells(1, TotalCols))
    ActiveSheet.UsedRange.AutoFilter 1, "=PO NUMBER", xlAnd
    Cells.Delete

    'Add column headers
    Rows(1).Insert
    Range(Cells(1, 1), Cells(UBound(ColHeaders), UBound(ColHeaders, 2))) = ColHeaders
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count
    
    'Check to see if all POs were found
    For i = 1 To UBound(POList)
        For j = 2 To TotalRows
            If Range("A" & j).Value = POList(i) Then
                Exit For
            Else
                If j = TotalRows Then
                    MsgBox "PO # " & POList(i) & " could not be found."
                End If
            End If
        Next
    Next
    
    'Create columns needed on report
    CreateReport

    'Move the cursor to the cell below the button that should be run next
    Sheets("Macro").Select
    Range("E14").Select

    Application.ScreenUpdating = True
    Exit Sub

Import_Error:
    If Err.Number = 13 Then
        MsgBox "You must enter a number", vbOKOnly, "Error"
        Resume
    ElseIf Err.Number = 6 Then
        MsgBox "The number entered was too large", vbOKOnly, "Error"
        Resume
    ElseIf Err.Number = Errors.USER_INTERRUPT Then
        MsgBox "Macro canceled by user."
        Clean
        Application.ScreenUpdating = True
        Exit Sub
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure 'Button1' of module 'Program'"
        Clean
        Application.ScreenUpdating = True
        Exit Sub
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : ShipLog
' Date : 12/27/2013
' Desc : Formats the shipping log and adds kit boms
'---------------------------------------------------------------------------------------
Sub CreateShipLog()
    Application.ScreenUpdating = False

    'Import kit components
    AddKitLines

    'Add formatting to the ship log
    FormatReport

    'Save and close
    MsgBox "Please save the shipping log to your computer.", vbOKOnly, "Macro Finished!"
    ExportReport
    ThisWorkbook.Close

    Application.ScreenUpdating = True
End Sub

'---------------------------------------------------------------------------------------
' Proc : RemoveLines
' Date : 12/27/2013
' Desc : Removes line items from the shipment
'---------------------------------------------------------------------------------------
Sub RemoveLines()
    Dim ColHeaders As Variant
    Dim TotalCols As Long
    Dim PO As String
    Dim StartLn As Long
    Dim EndLn As Long
    Dim i As Long

    Sheets("Ship Log").Select
    TotalCols = ActiveSheet.UsedRange.Columns.Count
    ColHeaders = Range(Cells(1, 1), Cells(1, TotalCols)).Value

    frmRemoveLines.Show

    PO = frmRemoveLines.PONumber
    StartLn = frmRemoveLines.StartLine
    EndLn = frmRemoveLines.EndLine
    
    Unload frmRemoveLines

    If PO <> 0 And StartLn <> 0 And EndLn <> 0 Then
        For i = StartLn To EndLn Step 10
            ActiveSheet.UsedRange.AutoFilter 2, "=" & PO & "/" & i
            Cells.Delete
            Rows(1).Insert
            Range(Cells(1, 1), Cells(1, TotalCols)).Value = ColHeaders
        Next
    End If

    Sheets("Macro").Select
    Range("G7").Select
End Sub

'---------------------------------------------------------------------------------------
' Proc : Clean
' Date : 8/29/2013
' Desc : Removes all data added at runtime
'---------------------------------------------------------------------------------------
Sub Clean()
    Dim PrevDispAlert As Boolean
    Dim PrevActiveBook As Workbook
    Dim s As Worksheet

    PrevDispAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False

    Set PrevActiveBook = ActiveWorkbook
    ThisWorkbook.Activate

    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.AutoFilterMode = False
            s.Cells.Delete
            s.Range("A1").Select
        End If
    Next

    Sheets("Macro").Select
    Range("C7").Select

    PrevActiveBook.Activate
    Application.DisplayAlerts = PrevDispAlert
End Sub

Attribute VB_Name = "Program"
Option Explicit
Public Const VersionNumber As String = "1.0.0"

Sub Button1()
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim ColHeaders() As Variant
    Dim NumOfPOs As Integer
    Dim RowCount As Long
    Dim PO As String
    Dim i As Long

    Application.ScreenUpdating = False
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
    NumOfPOs = CInt(InputBox("Number of POs on shipment", "PO Quantity"))
    On Error GoTo 0

    'Prompt user for PO numbers
    'Copy the PO that was entered onto Ship Log
    Sheets("POR").Select
    ReDim POList(1 To NumOfPOs) As String
    For i = 1 To NumOfPOs
        PO = InputBox("Enter PO #" & i, "PO Entry")
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

    'Create columns needed on report
    CreateReport

    'Move the cursor to the cell below the button that should be run next
    Sheets("Macro").Select
    Range("G7").Select

    'Go back to the sheet that may require user action
    Sheets("Ship Log").Select
    Range("A1").Select

    Application.ScreenUpdating = True

    MsgBox "1. Verify that all lines have a SIM/PART number." & vbCrLf & _
           "2. Go to the 'Macro' sheet and click 'Import Kit Lines'" & vbCrLf & vbCrLf & _
           "NOTE: If the item is a Club Car part please email the PART/SIM to TReische@wesco.com"

    Exit Sub

Import_Error:
    If Err.Number = 13 Then
        MsgBox "You must enter a number", vbOKOnly, "Error"
        Resume
    ElseIf Err.Number = 6 Then
        MsgBox "The number entered was too large", vbOKOnly, "Error"
        Resume
    Else
        MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure 'Button1' of module 'Program'"
        Clean
        Application.ScreenUpdating = True
        Exit Sub
    End If

End Sub

Sub Button2()
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

'Remove PO Lines
Sub Button3()
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

    If PO <> 0 And StartLn <> 0 And EndLn <> 0 Then
        For i = StartLn To EndLn Step 10
            ActiveSheet.UsedRange.AutoFilter 3, "=" & PO & "/" & i
            Cells.Delete
            Rows(1).Insert
            Range(Cells(1, 1), Cells(1, TotalCols)).Value = ColHeaders
        Next
    End If
    
    Sheets("Macro").Select
    Range("E14").Select
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
















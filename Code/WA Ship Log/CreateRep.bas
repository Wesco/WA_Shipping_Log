Attribute VB_Name = "CreateRep"
Option Explicit

Sub CreateReport()
    Dim Lookup As String
    Dim Part As Variant
    Dim Desc As Variant
    Dim TotalRows As Long
    Dim LookupTable As Variant
    Dim i As Long
    Dim j As Long

    Sheets("Ship Log").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'PO/LN
    Columns(1).Insert
    Range("A1").Value = "PO/LN"
    With Range("A2:A" & TotalRows)
        .Formula = "=RIGHT(""000000"" & B2,6) & ""/"" & C2"
        .NumberFormat = "@"
        .Value = .Value
    End With
    Columns("B:C").Delete

    'Description
    Columns(2).Insert
    Range("B1").Value = "Description"
    With Range("B2:B" & TotalRows)
        .NumberFormat = "General"
        .Formula = "=TRIM(C2)"
        .Value = .Value
    End With
    Range("C:C").Delete

    'SIMs
    Columns("B:C").Insert
    Range("B1").Value = "SIM"
    Range("C1").Value = "Part"
    Sheets("Master").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    LookupTable = Range(Cells(1, 1), Cells(TotalRows, 2))

    Sheets("Ship Log").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    Desc = Range("D2:D" & TotalRows)

    'Try to find item part numbers in the description field
    ' i = Current row on worksheet
    ' j = Current string in description
    For i = 2 To TotalRows
        'Desc starts at row 2
        'If there is only 1 row data Part is a single dimensional array,
        'otherwise it is multidimensional
        If TotalRows > 2 Then
            Part = Split(Desc(i - 1, 1), " ")
        Else
            Part = Split(Desc, " ")
        End If

        'Skip errors here because if the vlookup fails an error will be throw.
        On Error Resume Next
        For j = 0 To UBound(Part)
            If Lookup = "" Then
                Lookup = WorksheetFunction.VLookup(Part(j), LookupTable, 2, False)
            End If
            If Lookup <> "" Then
                Cells(i, 2).Value = Lookup
                Cells(i, 3).Value = Part(j)
                Lookup = ""
                Exit For
            End If
        Next
        On Error GoTo 0
    Next

    'Remove Account lines
    For i = TotalRows To 1 Step -1
        If InStr(Range("D" & i).Value, "ACCOUNT NO") Then
            Rows(i).Delete
        End If
    Next
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Ticket #
    Columns(1).Insert
    Range("A1").Value = "Ticket/LN"
    With Range("A2:A" & TotalRows)
        .Formula = "=IF(TRIM(RIGHT(""000000"" & H2,6) & ""/"" & I2)=""/"","""",RIGHT(""000000"" & H2,6) & ""/"" & I2)"
        .Value = .Value
    End With
    Columns("G:I").Delete

    'Qty Sent
    Range("F1").Value = "Qty Sent"

    'Kit Qty
    Range("F1:F" & TotalRows).Copy Range("G1")
    Range("G1").Value = "Kit Qty"
    
End Sub

Sub AddKitLines()
    Dim KitLookup As String
    Dim KitBOM() As String
    Dim LookupTable As Variant
    Dim MasterTable As Variant
    Dim GapsTable As Variant
    Dim StartRow As Long
    Dim KitLnCount As Integer
    Dim KitBOMRows As Long
    Dim CurrentSIM As String
    Dim CurrentLine As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim l As Long

    Sheets("Gaps").Select
    GapsTable = Range(Cells(1, 2), Cells(ActiveSheet.UsedRange.Rows.Count, 2))

    Sheets("Master").Select
    MasterTable = Range(Cells(1, 2), Cells(ActiveSheet.UsedRange.Rows.Count, 3))

    Sheets("Kit BOM").Select
    KitBOMRows = ActiveSheet.UsedRange.Rows.Count
    LookupTable = Range(Cells(1, 3), Cells(KitBOMRows, 3))

    Sheets("Ship Log").Select

    'Start at line 2 of Ship Log
    i = 1
    Do While i <= Sheets("Ship Log").UsedRange.Rows.Count
        i = i + 1
        KitLnCount = 0
        CurrentSIM = Sheets("Ship Log").Range("C" & i).Value

        'Check to see if the item is listed in the Kit BOM
        'If the item is not on the Kit BOM KitLookup is set to an empty string
        On Error GoTo NotFound
        KitLookup = WorksheetFunction.VLookup(CurrentSIM, LookupTable, 1, False)
        On Error GoTo 0

        If KitLookup <> "" Then
            'Get the starting row of the kit
            StartRow = CLng(WorksheetFunction.Match(CurrentSIM, LookupTable, 0))

            For j = StartRow To KitBOMRows
                If Sheets("Kit BOM").Range("C" & j).Value = KitLookup Then
                    KitLnCount = KitLnCount + 1
                Else
                    'Subtract 2 from kit line count because the first line is the Kit SIM
                    'and the last line is a special note denoting the end of the kit
                    ReDim KitBOM(1 To KitLnCount - 1, 1 To 7) As String

                    For k = 1 To KitLnCount - 1
                        CurrentLine = j - KitLnCount + k

                        'Ticket/LN
                        KitBOM(k, 1) = Sheets("Ship Log").Range("A" & i).Value
                        
                        'PO/LN
                        KitBOM(k, 2) = Sheets("Ship Log").Range("B" & i).Value

                        'SIM
                        KitBOM(k, 3) = Sheets("Kit BOM").Range("F" & CurrentLine).Value

                        'Part
                        On Error GoTo PartNotFound
                        KitBOM(k, 4) = WorksheetFunction.VLookup(KitBOM(k, 3), MasterTable, 2, False)
                        On Error GoTo 0

                        'Description
                        KitBOM(k, 5) = Sheets("Kit BOM").Range("I" & CurrentLine).Value

                        'Qty Sent
                        KitBOM(k, 6) = Sheets("Kit BOM").Range("G" & CurrentLine).Value * Sheets("Ship Log").Range("F" & i).Value

                        'Kit Qty, Left blank because only kit components are being added
                        KitBOM(k, 7) = ""
                    Next

                    If KitLnCount > 0 Then
                        'i + 1 = line below the current kit, the array holds the kits components
                        'Subtract 1 from kit line count because the first line is the Kit SIM
                        'and was not read into the array
                        Rows(i + 1 & ":" & i + KitLnCount - 1).Insert
                        Range(Cells(i + 1, 1), Cells(i + KitLnCount - 1, 7)).Value = KitBOM
                        
                        'Add the number of lines inserted to 'i' so that when it increments
                        'it will not select a kit component that was just inserted
                        i = i + KitLnCount - 1
                    End If
                    Exit For
                End If
            Next
        End If
    Loop

    'Qty Sent - Store as values since they were read in as strings from the array
    With Range(Cells(2, 6), Cells(ActiveSheet.UsedRange.Rows.Count, 6))
        .Value = .Value
    End With
    ActiveSheet.UsedRange.Columns.AutoFit
    Exit Sub


NotFound:
    KitLookup = ""
    Resume Next

PartNotFound:
    KitBOM(k, 6) = Sheets("Kit BOM").Range("H" & CurrentLine)
    Resume Next
End Sub

Sub FormatReport()
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim ColKitQty As Integer
    Dim i As Long
    Dim j As Long

    Sheets("Ship Log").Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = ActiveSheet.UsedRange.Columns.Count
    ColKitQty = FindColumn("Kit Qty")

    'Skid # - Notes: Wesco to Workability = Blue
    Range("A1:G1").Interior.Color = 12419407

    With Range("A1:G1")
        .Font.Color = rgbWhite
        .Font.Bold = True
    End With

    'Insert a blank row between each kit and alternate
    'kit colors between light blue / light purple
    ' i - Keeps track of current row
    ' j - Keeps track of current kit
    Do While i < TotalRows + j
        i = i + 1
        If Cells(i, ColKitQty).Value <> "" And i > 2 Then
            Rows(i).Insert
            Range(Cells(i, 1), Cells(i, 7)).ClearFormats
            i = i + 1
            j = j + 1
        End If

        If j Mod 2 = 1 Then
            'Light blue
            Range(Cells(i, 1), Cells(i, 7)).Interior.Color = 14994616
        ElseIf i > 1 Then
            'Light purple
            Range(Cells(i, 1), Cells(i, 7)).Interior.Color = 14336204
        End If
    Loop

    ActiveSheet.UsedRange.Columns.AutoFit
End Sub

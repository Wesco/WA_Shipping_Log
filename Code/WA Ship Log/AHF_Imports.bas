Attribute VB_Name = "AHF_Imports"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc  : Sub ImportGaps
' Date  : 12/12/2012
' Desc  : Imports gaps to the workbook containing this macro.
' Ex    : ImportGaps
'---------------------------------------------------------------------------------------
Sub ImportGaps()
    Dim sPath As String     'Gaps file path
    Dim sName As String     'Gaps Sheet Name
    Dim iCounter As Long    'Counter to decrement the date
    Dim iRows As Long       'Total number of rows
    Dim dt As Date          'Date for gaps file name and path
    Dim Result As VbMsgBoxResult    'Yes/No to proceed with old gaps file if current one isn't found
    Dim Gaps As Worksheet           'The sheet named gaps if it exists, else this = nothing
    Dim StartTime As Double         'The time this function was started
    Dim FileFound As Boolean        'Indicates whether or not gaps was found

    StartTime = Timer
    FileFound = False

    'This error is bypassed so you can determine whether or not the sheet exists
    On Error GoTo CREATE_GAPS
    Set Gaps = ThisWorkbook.Sheets("Gaps")
    On Error GoTo 0

    Application.DisplayAlerts = False

    'Find gaps
    For iCounter = 0 To 15
        dt = Date - iCounter
        sPath = "\\br3615gaps\gaps\3615 Gaps Download\" & Format(dt, "yyyy") & "\"
        sName = "3615 " & Format(dt, "yyyy-mm-dd") & ".csv"
        If FileExists(sPath & sName) Then
            FileFound = True
            Exit For
        End If
    Next

    'Make sure Gaps file was found
    If FileFound = True Then
        If dt <> Date Then
            Result = MsgBox( _
                     Prompt:="Gaps from " & Format(dt, "mmm dd, yyyy") & " was found." & vbCrLf & "Would you like to continue?", _
                     Buttons:=vbYesNo, _
                     Title:="Gaps not up to date")
        End If

        If Result <> vbNo Then
            If ThisWorkbook.Sheets("Gaps").Range("A1").Value <> "" Then
                Gaps.Cells.Delete
            End If

            Workbooks.Open sPath & sName
            ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Gaps").Range("A1")
            ActiveWorkbook.Close

            Sheets("Gaps").Select
            iRows = ActiveSheet.UsedRange.Rows.Count
            Columns(1).EntireColumn.Insert
            Range("A1").Value = "SIM"
            Range("A2").Formula = "=C2&D2"
            Range("A2").AutoFill Destination:=Range(Cells(2, 1), Cells(iRows, 1))
            Range(Cells(2, 1), Cells(iRows, 1)).Value = Range(Cells(2, 1), Cells(iRows, 1)).Value
        Else
            Err.Raise 18, "ImportGaps", "Import canceled"
        End If
    Else
        Err.Raise 53, "ImportGaps", "Gaps could not be found."
    End If

    Application.DisplayAlerts = True
    Exit Sub

CREATE_GAPS:
    ThisWorkbook.Sheets.Add After:=Sheets(ThisWorkbook.Sheets.Count)
    ActiveSheet.Name = "Gaps"
    Resume

End Sub

'---------------------------------------------------------------------------------------
' Proc : UserImportFile
' Date : 1/29/2013
' Desc : Prompts the user to select a file for import
'---------------------------------------------------------------------------------------
Sub UserImportFile(DestRange As Range, Optional DelFile As Boolean = False, Optional ShowAllData As Boolean = False, Optional SourceSheet As String = "")
    Dim File As String              'Full path to user selected file
    Dim FileDate As String          'Date the file was last modified
    Dim OldDispAlert As Boolean     'Original state of Application.DisplayAlerts

    OldDispAlert = Application.DisplayAlerts
    File = Application.GetOpenFilename()

    Application.DisplayAlerts = False
    If File <> "False" Then
        FileDate = Format(FileDateTime(File), "mm/dd/yy")
        Workbooks.Open File
        If SourceSheet = "" Then SourceSheet = ActiveSheet.Name
        If ShowAllData = True Then
            On Error Resume Next
            ActiveSheet.AutoFilter.ShowAllData
            ActiveSheet.UsedRange.Columns.Hidden = False
            ActiveSheet.UsedRange.Rows.Hidden = False
            On Error GoTo 0
        End If
        Sheets(SourceSheet).UsedRange.Copy Destination:=DestRange
        ActiveWorkbook.Close
        ThisWorkbook.Activate

        If DelFile = True Then
            DeleteFile File
        End If
    Else
        Err.Raise 18
    End If
    Application.DisplayAlerts = OldDispAlert
End Sub

'---------------------------------------------------------------------------------------
' Proc : Import473
' Date : 4/11/2013
' Desc : Imports a 473 report from the current day
'---------------------------------------------------------------------------------------
Sub Import473(Destination As Range, Optional Branch As String = "3615")
    Dim sPath As String
    Dim FileName As String
    Dim AlertStatus As Boolean

    FileName = "473 " & Format(Date, "m-dd-yy") & ".xlsx"
    sPath = "\\br3615gaps\gaps\" & Branch & " 473 Download\" & FileName
    AlertStatus = Application.DisplayAlerts

    If FileExists(sPath) Then
        Workbooks.Open sPath
        ActiveSheet.UsedRange.Copy Destination:=Destination

        Application.DisplayAlerts = False
        ActiveWorkbook.Close
        Application.DisplayAlerts = AlertStatus
    Else
        MsgBox Prompt:="473 report not found."
        Err.Raise 18
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : ImportSupplierContacts
' Date : 4/22/2013
' Desc : Imports the supplier contact master list
'---------------------------------------------------------------------------------------
Sub ImportSupplierContacts(Destination As Range)
    Const sPath As String = "\\br3615gaps\gaps\Contacts\Supplier Contact Master.xlsx"
    Dim PrevDispAlerts As Boolean

    PrevDispAlerts = Application.DisplayAlerts

    Workbooks.Open sPath
    ActiveSheet.UsedRange.Copy Destination:=Destination

    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Application.DisplayAlerts = PrevDispAlerts
End Sub

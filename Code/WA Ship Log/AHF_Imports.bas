Attribute VB_Name = "AHF_Imports"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc  : Sub ImportGaps
' Date  : 12/12/2012
' Desc  : Imports gaps to the workbook containing this macro.
' Ex    : ImportGaps
'---------------------------------------------------------------------------------------
Sub ImportGaps(Optional Destination As Range, Optional SimsAsText As Boolean = True)
    Dim Path As String      'Gaps file path
    Dim Name As String      'Gaps Sheet Name
    Dim i As Long           'Counter to decrement the date
    Dim dt As Date          'Date for gaps file name and path
    Dim TotalRows As Long   'Total number of rows
    Dim Result As VbMsgBoxResult    'Yes/No to proceed with old gaps file if current one isn't found


    'This error is bypassed so you can determine whether or not the sheet exists
    On Error GoTo CREATE_GAPS
    If TypeName(Destination) = "Nothing" Then
        Set Destination = ThisWorkbook.Sheets("Gaps").Range("A1")
    End If
    On Error GoTo 0

    Application.DisplayAlerts = False

    'Try to find Gaps
    For i = 0 To 15
        dt = Date - i
        Path = "\\br3615gaps\gaps\3615 Gaps Download\" & Format(dt, "yyyy") & "\"
        Name = "3615 " & Format(dt, "yyyy-mm-dd") & ".csv"
        If Exists(Path & Name) Then
            Exit For
        End If
    Next

    'Make sure Gaps file was found
    If Exists(Path & Name) Then
        If dt <> Date Then
            Result = MsgBox( _
                     Prompt:="Gaps from " & Format(dt, "mmm dd, yyyy") & " was found." & vbCrLf & "Would you like to continue?", _
                     Buttons:=vbYesNo, _
                     Title:="Gaps not up to date")
        End If

        If Result <> vbNo Then
            ThisWorkbook.Activate
            Sheets(Destination.Parent.Name).Select

            'If there is data on the destination sheet delete it
            If Range("A1").Value <> "" Then
                Cells.Delete
            End If

            ImportCsvAsText Path, Name, Sheets("Gaps").Range("A1")
            TotalRows = ActiveSheet.UsedRange.Rows.Count
            Columns(1).Insert
            Range("A1").Value = "SIM"

            'SIMs are 11 digits and can have leading 0's
            If SimsAsText = True Then
                Range("A2:A" & TotalRows).Formula = "=""=""&""""""""&RIGHT(""000000"" & C2, 6)&RIGHT(""00000"" & D2, 5)&"""""""""
            Else
                Range("A2:A" & TotalRows).Formula = "=C2&RIGHT(""00000"" & D2, 5)"
            End If

            Range("A2:A" & TotalRows).Value = Range("A2:A" & TotalRows).Value
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
' Proc : ImportCsvAsText
' Date : 7/1/2014
' Desc : Import a CSV file with all fields as text
'---------------------------------------------------------------------------------------
Sub ImportCsvAsText(Path As String, File As String, Destination As Range)
    Dim Name As String
    Dim FileNo As Integer
    Dim TotalCols As Long
    Dim ColHeaders As String
    Dim ColFormat As Variant
    Dim i As Long


    'Make sure path ends with a trailing slash
    If Right(Path, 1) <> "\" Then Path = Path & "\"

    'If the file exists open it
    If FileExists(Path & File) Then
        Name = Left(File, InStrRev(File, ".") - 1)

        'Read first line of file to figure out how many columns there are
        FileNo = FreeFile()
        Open Path & File For Input As #FileNo
        Line Input #FileNo, ColHeaders
        Close #FileNo

        TotalCols = UBound(Split(ColHeaders, ",")) + 1

        'Set column format to 2 (text) for all columns
        ReDim ColFormat(1 To TotalCols)
        For i = 1 To TotalCols
            ColFormat(i) = 2
        Next

        'Import CSV
        With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & Path & File, Destination:=Destination)
            .Name = Name
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 437
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = ColFormat
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
        End With

        'Remove the connection
        ActiveWorkbook.Connections(Name).Delete
        ActiveSheet.QueryTables(ActiveSheet.QueryTables.Count).Delete
    Else
        Err.Raise 53, "OpenCsvAsText", "File not found"
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : UserImportFile
' Date : 1/29/2013
' Desc : Prompts the user to select a file for import
'---------------------------------------------------------------------------------------
Sub UserImportFile(DestRange As Range, Optional DelFile As Boolean = False, Optional ShowAllData As Boolean = False, Optional SourceSheet As String = "", Optional FileFilter = "")
    Dim File As String              'Full path to user selected file
    Dim FileDate As String          'Date the file was last modified
    Dim OldDispAlert As Boolean     'Original state of Application.DisplayAlerts

    OldDispAlert = Application.DisplayAlerts
    File = Application.GetOpenFilename(FileFilter)

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
' Proc : Import117byISN
' Date : 4/10/2013
' Desc : Imports the most recent 117 report for the specified sales number
'---------------------------------------------------------------------------------------
Sub Import117byISN(RepType As ReportType, Destination As Range, Optional ByVal ISN As String = "", Optional Cancel As Boolean = False)
    Dim sPath As String
    Dim FileName As String

    If ISN = "" And Cancel = False Then
        ISN = InputBox("Inside Sales Number:", "Please enter the ISN#")
    Else
        If ISN = "" Then
            Err.Raise 53
        End If
    End If

    If ISN <> "" Then
        Select Case RepType
            Case ReportType.DS:
                FileName = "3615 " & Format(Date, "m-dd-yy") & " DSORDERS.xlsx"

            Case ReportType.BO:
                FileName = "3615 " & Format(Date, "m-dd-yy") & " BACKORDERS.xlsx"

            Case ReportType.ALL
                FileName = "3615 " & Format(Date, "m-dd-yy") & " ALLORDERS.xlsx"
        End Select

        sPath = "\\br3615gaps\gaps\3615 117 Report\ByInsideSalesNumber\" & ISN & "\" & FileName

        If Exists(sPath) Then
            Workbooks.Open sPath
            ActiveSheet.UsedRange.Copy Destination:=Destination
            Application.DisplayAlerts = False
            ActiveWorkbook.Close
            Application.DisplayAlerts = True
        Else
            MsgBox Prompt:=ReportTypeText(RepType) & " report not found.", Title:="Error 53"
        End If
    Else
        Err.Raise 18
    End If
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

    If Exists(sPath) Then
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

'---------------------------------------------------------------------------------------
' Proc  : Function Exists
' Date  : 6/24/14
' Type  : Boolean
' Desc  : Checks if a file exists and can be read
' Ex    : FileExists "C:\autoexec.bat"
'---------------------------------------------------------------------------------------
Private Function Exists(ByVal FilePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    'Remove trailing backslash
    If InStr(Len(FilePath), FilePath, "\") > 0 Then
        FilePath = Left(FilePath, Len(FilePath) - 1)
    End If

    'Check to see if the file exists and has read access
    On Error GoTo File_Error
    If fso.FileExists(FilePath) Then
        fso.OpenTextFile(FilePath, 1).Read 0
        Exists = True
    Else
        Exists = False
    End If
    On Error GoTo 0

    Exit Function

File_Error:
    Exists = False
End Function

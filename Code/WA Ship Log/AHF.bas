Attribute VB_Name = "AHF"
Option Explicit

'List of error codes
Enum Errors
    USER_INTERRUPT = 18
    FILE_NOT_FOUND = 53
    FILE_ALREADY_OPEN = 55
    FILE_ALREADY_EXISTS = 58
    DISK_FULL = 63
    PERMISSION_DENIED = 70
    PATH_FILE_ACCESS_ERROR = 75
    PATH_NOT_FOUND = 76
    ORBJECT_OR_WITH_BLOCK_NOT_SET = 91
    INVALID_FILE_FORMAT = 321
    OUT_OF_MEMORY = 31001
    ERROR_SAVING_FILE = 31036
    ERROR_LOADING_FROM_FILE = 31037
End Enum

'List of custom error messages
Enum CustErr
    COLNOTFOUND = 50000
End Enum

'Used when importing 117 to determine the type of report to pull
Enum ReportType
    DS
    BO
    ALL
    INQ
End Enum

'---------------------------------------------------------------------------------------
' Proc : FilterSheet
' Date : 1/29/2013
' Desc : Remove all rows that do not match a specified string
'---------------------------------------------------------------------------------------
Sub FilterSheet(sFilter As String, ColNum As Integer, Match As Boolean)
    Dim Rng As Range
    Dim aRng() As Variant
    Dim aHeaders As Variant
    Dim StartTime As Double
    Dim iCounter As Long
    Dim i As Long
    Dim y As Long

    StartTime = Timer
    Set Rng = ActiveSheet.UsedRange
    aHeaders = Range(Cells(1, 1), Cells(1, ActiveSheet.UsedRange.Columns.Count))
    iCounter = 1

    Do While iCounter <= Rng.Rows.Count
        If Match = True Then
            If Rng(iCounter, ColNum).Value = sFilter Then
                i = i + 1
            End If
        Else
            If Rng(iCounter, ColNum).Value <> sFilter Then
                i = i + 1
            End If
        End If
        iCounter = iCounter + 1
    Loop

    ReDim aRng(1 To i, 1 To Rng.Columns.Count) As Variant

    iCounter = 1
    i = 0
    Do While iCounter <= Rng.Rows.Count
        If Match = True Then
            If Rng(iCounter, ColNum).Value = sFilter Then
                i = i + 1
                For y = 1 To Rng.Columns.Count
                    aRng(i, y) = Rng(iCounter, y)
                Next
            End If
        Else
            If Rng(iCounter, ColNum).Value <> sFilter Then
                i = i + 1
                For y = 1 To Rng.Columns.Count
                    aRng(i, y) = Rng(iCounter, y)
                Next
            End If
        End If
        iCounter = iCounter + 1
    Loop

    ActiveSheet.Cells.Delete
    Range(Cells(1, 1), Cells(UBound(aRng, 1), UBound(aRng, 2))) = aRng
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, UBound(aHeaders, 2))) = aHeaders
End Sub

'---------------------------------------------------------------------------------------
' Proc : ExportCode
' Date : 3/19/2013
' Desc : Exports all modules
'---------------------------------------------------------------------------------------
Sub ExportCode()
    Dim comp As Variant
    Dim codeFolder As String
    Dim FileName As String
    Dim File As String

    'References Microsoft Visual Basic for Applications Extensibility 5.3
    AddReference "{0002E157-0000-0000-C000-000000000046}", 5, 3
    codeFolder = GetWorkbookPath & "Code\" & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & "\"

    On Error Resume Next
    RecMkDir codeFolder
    On Error GoTo 0

    'Remove all previously exported modules
    File = Dir(codeFolder)
    Do While File <> ""
        DeleteFile codeFolder & File
        File = Dir
    Loop

    'Export modules in current project
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1
                FileName = codeFolder & comp.Name & ".bas"
                comp.Export FileName
            Case 2
                FileName = codeFolder & comp.Name & ".cls"
                comp.Export FileName
            Case 3
                FileName = codeFolder & comp.Name & ".frm"
                comp.Export FileName
        End Select
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : ImportModule
' Date : 4/4/2013
' Desc : Imports a code module into VBProject
'---------------------------------------------------------------------------------------
Sub ImportModule()
    Dim comp As Variant
    Dim codeFolder As String
    Dim FileName As String
    Dim WkbkPath As String

    'Adds a reference to Microsoft Visual Basic for Applications Extensibility 5.3
    AddReference "{0002E157-0000-0000-C000-000000000046}", 5, 3

    'Gets the path to this workbook
    WkbkPath = Left$(ThisWorkbook.fullName, InStr(1, ThisWorkbook.fullName, ThisWorkbook.Name, vbTextCompare) - 1)

    'Gets the path to this workbooks code
    codeFolder = WkbkPath & "Code\" & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & "\"

    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Name <> "All_Helper_Functions" Then
            Select Case comp.Type
                Case 1
                    FileName = codeFolder & comp.Name & ".bas"
                    ThisWorkbook.VBProject.VBComponents.Remove comp
                    ThisWorkbook.VBProject.VBComponents.Import FileName
                Case 2
                    FileName = codeFolder & comp.Name & ".cls"
                    ThisWorkbook.VBProject.VBComponents.Remove comp
                    ThisWorkbook.VBProject.VBComponents.Import FileName
                Case 3
                    FileName = codeFolder & comp.Name & ".frm"
                    ThisWorkbook.VBProject.VBComponents.Remove comp
                    ThisWorkbook.VBProject.VBComponents.Import FileName
            End Select
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : GetWorkbookPath
' Date : 3/19/2013
' Desc : Gets the full path of ThisWorkbook
'---------------------------------------------------------------------------------------
Function GetWorkbookPath() As String
    Dim fullName As String
    Dim wrkbookName As String
    Dim pos As Long

    wrkbookName = ThisWorkbook.Name
    fullName = ThisWorkbook.fullName

    pos = InStr(1, fullName, wrkbookName, vbTextCompare)

    GetWorkbookPath = Left$(fullName, pos - 1)
End Function

'---------------------------------------------------------------------------------------
' Proc : EndsWith
' Date : 3/19/2013
' Desc : Checks if a string ends in a specified character
'---------------------------------------------------------------------------------------
Function EndsWith(ByVal InString As String, ByVal TestString As String) As Boolean
    EndsWith = (Right$(InString, Len(TestString)) = TestString)
End Function

'---------------------------------------------------------------------------------------
' Proc : AddReferences
' Date : 3/19/2013
' Desc : Adds a reference to VBProject
'---------------------------------------------------------------------------------------
Sub AddReference(GUID As String, Major As Integer, Minor As Integer)
    Dim ID As Variant
    Dim Ref As Variant
    Dim Result As Boolean


    For Each Ref In ThisWorkbook.VBProject.References
        If Ref.GUID = GUID And Ref.Major = Major And Ref.Minor = Minor Then
            Result = True
        End If
    Next

    'References Microsoft Visual Basic for Applications Extensibility 5.3
    If Result = False Then
        ThisWorkbook.VBProject.References.AddFromGuid GUID, Major, Minor
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : RemoveReferences
' Date : 3/19/2013
' Desc : Removes a reference from VBProject
'---------------------------------------------------------------------------------------
Sub RemoveReference(GUID As String, Major As Integer, Minor As Integer)
    Dim Ref As Variant

    For Each Ref In ThisWorkbook.VBProject.References
        If Ref.GUID = GUID And Ref.Major = Major And Ref.Minor = Minor Then
            Application.VBE.ActiveVBProject.References.Remove Ref
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : ShowReferences
' Date : 4/4/2013
' Desc : Lists all VBProject references
'---------------------------------------------------------------------------------------
Sub ShowReferences()
    Dim i As Variant
    Dim n As Integer

    ThisWorkbook.Activate
    On Error GoTo SHEET_EXISTS
    Sheets("VBA References").Select
    ActiveSheet.Cells.Delete
    On Error GoTo 0

    [A1].Value = "Name"
    [B1].Value = "Description"
    [C1].Value = "GUID"
    [D1].Value = "Major"
    [E1].Value = "Minor"

    For i = 1 To ThisWorkbook.VBProject.References.Count
        n = i + 1
        With ThisWorkbook.VBProject.References(i)
            Cells(n, 1).Value = .Name
            Cells(n, 2).Value = .Description
            Cells(n, 3).Value = .GUID
            Cells(n, 4).Value = .Major
            Cells(n, 5).Value = .Minor
        End With
    Next
    Columns.AutoFit

    Exit Sub

SHEET_EXISTS:
    ThisWorkbook.Sheets.Add After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count), Count:=1
    ActiveSheet.Name = "VBA References"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : ReportTypeText
' Date : 4/10/2013
' Desc : Returns the report type as a string
'---------------------------------------------------------------------------------------
Function ReportTypeText(RepType As ReportType) As String
    Select Case RepType
        Case ReportType.BO:
            ReportTypeText = "BO"
        Case ReportType.DS:
            ReportTypeText = "DS"
        Case ReportType.ALL:
            ReportTypeText = "ALL"
        Case ReportType.INQ:
            ReportTypeText = "INQ"
    End Select
End Function

'---------------------------------------------------------------------------------------
' Proc : DeleteColumn
' Date : 4/11/2013
' Desc : Removes a column based on text in the column header
'---------------------------------------------------------------------------------------
Sub DeleteColumn(HeaderText As String)
    Dim i As Integer

    For i = ActiveSheet.UsedRange.Columns.Count To 1 Step -1
        If Trim(Cells(1, i).Value) = HeaderText Then
            Columns(i).Delete
            Exit For
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : FindColumn
' Date : 4/11/2013
' Desc : Returns the column number if a match is found
'---------------------------------------------------------------------------------------
Function FindColumn(ByVal HeaderText As String, Optional SearchArea As Range) As Integer
    Dim i As Integer: i = 0
    Dim ColText As String

    If TypeName(SearchArea) = "Nothing" Or TypeName(SearchArea) = Empty Then
        Set SearchArea = Range(Cells(1, 1), Cells(1, ActiveSheet.UsedRange.Columns.Count))
    End If

    For i = 1 To SearchArea.Columns.Count
        ColText = Trim(SearchArea.Cells(1, i).Value)

        Do While InStr(ColText, "  ")
            ColText = Replace(ColText, "  ", " ")
        Loop

        If ColText = HeaderText Then
            FindColumn = i
            Exit For
        End If
    Next

    If FindColumn = 0 Then Err.Raise CustErr.COLNOTFOUND, "FindColumn", HeaderText
End Function

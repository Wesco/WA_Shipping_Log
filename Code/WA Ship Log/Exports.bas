Attribute VB_Name = "Exports"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : ExportReport
' Date : 8/29/2013
' Desc : Exports the shipping log to a new workbook and prompts the user to save it
'---------------------------------------------------------------------------------------
Sub ExportReport()
    Sheets("Ship Log").Copy
    
    Range("A1").Select
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True

    Application.Dialogs(xlDialogSaveAs).Show
End Sub

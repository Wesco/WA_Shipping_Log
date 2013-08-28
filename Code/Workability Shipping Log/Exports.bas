Attribute VB_Name = "Exports"
Option Explicit

Sub ExportReport()
    Sheets("Ship Log").Copy
    
    Range("A1").Select
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True

    Application.Dialogs(xlDialogSaveAs).Show
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRemoveLines 
   Caption         =   "Remove PO Lines"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3450
   OleObjectBlob   =   "frmRemoveLines.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRemoveLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PONumber As Long
Public StartLine As Long
Public EndLine As Long

Private Sub UserForm_Initialize()
    PONumber = 0
    StartLine = 0
    EndLine = 0
    txtPO.Text = ""
    txtStartLn.Text = ""
    txtEndLn.Text = ""
    txtPO.SetFocus
End Sub

Private Sub btnRemove_Click()
    On Error GoTo Type_Err
    PONumber = CLng(txtPO)
    StartLine = CLng(txtStartLn)
    EndLine = CLng(txtEndLn)
    On Error GoTo 0

    If StartLine > EndLine Then
        MsgBox "The start line cannot be greater than the end line."
    ElseIf StartLine <= 0 Then
        MsgBox "The start line can not be less than or equal to zero."
    ElseIf PONumber <= 0 Then
        MsgBox "The PO number can not be less than or equal to zero."
    ElseIf Not StartLine Mod 10 = 0 Then
        MsgBox "The start line must be a multiple of ten."
    ElseIf Not EndLine Mod 10 = 0 Then
        MsgBox "The end line must be a multiple of ten."
    Else
        frmRemoveLines.Hide
    End If

    Exit Sub

Type_Err:
    If Err.Number = 13 Then
        MsgBox "Please make sure only numbers have been entered."
    Else
        Err.Raise Err.Number
    End If
End Sub

Private Sub btnCancel_Click()
    PONumber = 0
    StartLine = 0
    EndLine = 0
    frmRemoveLines.Hide
End Sub

Private Sub UserForm_Terminate()
    PONumber = 0
    StartLine = 0
    EndLine = 0
End Sub

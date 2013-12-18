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
    txtPO.Text = ""
    txtStartLn.Text = ""
    txtEndLn.Text = ""
    txtPO.SetFocus
End Sub

Private Sub btnRemove_Click()
    On Error GoTo Type_Err
    PONumber = CInt(txtPO)
    StartLine = CInt(txtStartLn)
    EndLine = CInt(txtEndLn)
    On Error GoTo 0

    If StartLine > EndLine Then
        MsgBox "The start line cannot be greater than the end line."
    ElseIf StartLine <= 0 Then
        MsgBox "The start line can not be less than or equal to zero."
    ElseIf PONumber <= 0 Then
        MsgBox "The PO number can not be less than or equal to zero."
    Else
        frmRemoveLines.Hide
    End If

    Exit Sub

Type_Err:
    If Err.Number = 13 Then
        MsgBox "Please make sure only numbers have been entered."
    End If
End Sub

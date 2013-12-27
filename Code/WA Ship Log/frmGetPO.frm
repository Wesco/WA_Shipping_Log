VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGetPO 
   Caption         =   "PO Entry"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3270
   OleObjectBlob   =   "frmGetPO.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGetPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PO As Long

Private Sub UserForm_Initialize()
    lblPONum.Caption = "Enter PO# " & POCount
    txtPO.Value = ""
    txtPO.SetFocus
    PO = 0
End Sub

Private Sub btnOk_Click()
    If txtPO.Value = "" Then
        MsgBox "You must enter a number."
        txtPO.SetFocus
    ElseIf txtPO.Value <= 0 Then
        MsgBox "You must enter a value greater than zero."
        txtPO.Text = ""
        txtPO.SetFocus
    Else
        On Error GoTo NumPO_Err
        PO = CLng(txtPO.Value)
        On Error GoTo 0
        frmGetPO.Hide
        txtPO.Text = ""
        txtPO.SetFocus
    End If

    Exit Sub

NumPO_Err:
    If Err.Number = 13 Then
        MsgBox "You must enter a number."
        txtPO.Text = ""
        txtPO.SetFocus
    ElseIf Err.Number = 6 Then
        MsgBox "The number you entered was too large."
        txtPO.Text = ""
        txtPO.SetFocus
    End If
End Sub

Private Sub btnCancel_Click()
    frmGetPO.Hide
End Sub

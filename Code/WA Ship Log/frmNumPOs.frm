VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmNumPOs 
   Caption         =   "PO Quantity"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4830
   OleObjectBlob   =   "frmNumPOs.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmNumPOs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public NumPOs As Integer

Private Sub UserForm_Initialize()
    NumPOs = 0
    txtNumPOs.SetFocus
    txtNumPOs.Text = ""
End Sub

Private Sub btnCancel_Click()
    NumPOs = 0
    txtNumPOs.Text = ""
    frmNumPOs.Hide
End Sub

Private Sub btnOk_Click()
    If txtNumPOs = "" Then
        MsgBox "You must enter a number."
        txtNumPOs.SetFocus
    Else
        On Error GoTo NumPO_Err
        NumPOs = CInt(txtNumPOs.Value)
        On Error GoTo 0
        frmNumPOs.Hide
    End If

    Exit Sub

NumPO_Err:
    If Err.Number = 13 Then
        MsgBox "You must enter a number."
        txtNumPOs.Text = ""
        txtNumPOs.SetFocus
    ElseIf Err.Number = 6 Then
        MsgBox "The number you entered was too large."
        txtNumPOs.Text = ""
        txtNumPOs.SetFocus
    End If
End Sub
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

Private Sub UserForm_Initialize()
    lblPONum.Caption = "Enter PO# " & NumOfPOs
    txtPO.Text = ""
    txtPO.SetFocus
End Sub

Private Sub btnOk_Click()

End Sub

Private Sub UserForm_Click()

End Sub

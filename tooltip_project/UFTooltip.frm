VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFTooltip 
   Caption         =   "MAIN UF"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UFTooltip.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub UserForm_Click()
    Debug.Print "TOOLTIP clicked!"
    Me.Hide
End Sub

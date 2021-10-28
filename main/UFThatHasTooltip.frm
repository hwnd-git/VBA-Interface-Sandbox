VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFThatHasTooltip 
   Caption         =   "TOOLTIP"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UFThatHasTooltip.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFThatHasTooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Implements ICanShowTooltip

Dim myTooltip As tooltip_project.UFTooltip

Private Property Set ICanShowTooltip_myTooltip(ByVal tooltipInstance As tooltip_project.UFTooltip)
    Set myTooltip = tooltipInstance
End Property

Private Property Get ICanShowTooltip_myTooltip() As tooltip_project.UFTooltip
    Set ICanShowTooltip_myTooltip = myTooltip
End Property

Private Sub ICanShowTooltip_showTooltip()
    Call myTooltip.Show
End Sub

Private Sub UserForm_Click()
    Debug.Print "MAIN UF clicked!"
    Call ICanShowTooltip_showTooltip
End Sub


Private Sub UserForm_Initialize()
    Set myTooltip = tooltip_project.New_UFTooltip
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} QEXimport 
   Caption         =   "QEX Import"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   OleObjectBlob   =   "QEXimport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "QEXimport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.width) - (0.5 * Me.width)
    Me.Top = Application.Top + (0.5 * Application.height) - (0.5 * Me.height)
End Sub

Private Sub Overwrite_Click()
    overwriteQTO = True
    Unload QEXimport
End Sub

Private Sub Append_Click()
    overwriteQTO = False
    Unload QEXimport
End Sub

Private Sub Cancel_Click()
    Unload QEXimport
    End
End Sub


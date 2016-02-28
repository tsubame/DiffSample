VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputAddress 
   Caption         =   "ADDRESS INPUT"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3180
   OleObjectBlob   =   "InputAddress.frx":0000
   StartUpPosition =   2  '‰æ–Ê‚Ì’†‰›
End
Attribute VB_Name = "InputAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

'Modul NAME     : InputAddress
'AUTHOR         : N.Togo

Private Sub CBOK_Click()

    If (Xadd.Text = "") Or (Yadd.Text = "") Then
        MsgBox "Please Input X,Y Address."
        Exit Sub
    End If
    
    TenkenX = Xadd.Text
    TenkenY = Yadd.Text
    
    Unload Me

End Sub

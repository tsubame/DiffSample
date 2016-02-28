VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WaferTenkenForm 
   Caption         =   "ADDRESS SET"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   OleObjectBlob   =   "WaferTenkenForm.frx":0000
   StartUpPosition =   2  '‰æ–Ê‚Ì’†‰›
End
Attribute VB_Name = "WaferTenkenForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Explicit

'Modul NAME     : WaferTENKEN
'AUTHOR         : N.Togo

Private Sub CBMove_Click()
    
    Dim TargetX As Integer
    Dim TargetY As Integer
    
    If (Xadd.Text = "") Or (Yadd.Text = "") Then
        MsgBox "Please Input X,Y Address."
        Exit Sub
    End If
    
    TargetX = Xadd.Text
    TargetY = Yadd.Text
    
    Unload Me
    
    Call XYMove(TargetX, TargetY)
        
End Sub
Private Sub CBClose_Click()
    Unload Me
End Sub



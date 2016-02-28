VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StringInputForm 
   Caption         =   "PARAMETER"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   OleObjectBlob   =   "StringInputForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "StringInputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()

    Unload Me
    
End Sub

Private Sub OKButton_Click()

    StringValue = TextBox1.Text
    Unload Me
    
End Sub

Private Sub UserForm_Activate()
    
    TextBox1.SetFocus

End Sub

Private Sub UserForm_Initialize()
    
    TextBox1.Text = ""
    
End Sub

Public Function Setup(Label As String) As String
    
    Label1.Caption = Label
    
End Function

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6585
   OleObjectBlob   =   "Menu.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
Private Sub openFileButton_Click()
    Application.Visible = False


    
    Dim path As Variant
        
    path = Application.GetOpenFilename

    Call ModuleExporter.showModules(path)

End Sub

Private Sub UserForm_Terminate()
    'MsgBox "Excel�̉�ʂ�\�����܂�"
    Application.Visible = True
End Sub
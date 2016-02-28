VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DcScenarioLoopOptionForm 
   Caption         =   "DC Test Scenario Looping Option"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7395
   OleObjectBlob   =   "DcScenarioLoopOptionForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "DcScenarioLoopOptionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'リストボックスのmultiSelectを2-fmMultiSelectExtendedに変更
'フォルダ選択ボタンのテキストを半角ピリオドに変更
'カテゴリアイテム移動ボタンのレイアウト配置を変更
'Form起動時のフォルダパスを、このブックのパス(＝JOBファイルのPath)に指定



Option Explicit
Public Event QueryClose(Cancel As Integer, CloseMode As Integer)

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    RaiseEvent QueryClose(Cancel, CloseMode)
End Sub

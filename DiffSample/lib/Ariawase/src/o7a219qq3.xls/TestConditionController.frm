VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TestConditionController 
   Caption         =   "TestConditionController"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   OleObjectBlob   =   "TestConditionController.frx":0000
End
Attribute VB_Name = "TestConditionController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Option Explicit


Public Event QueryClose(ByRef Cancel As Integer, ByVal CloseMode As Integer)   'TestConditionControllerフォーム終了告知イベント

Private Sub AbortButton_Click()
'Abortボタンが押下された時の処理
    'ここでは何もしない
    'CTestConditionControllerクラスでこのイベントを取得し、そこで処理する
    
End Sub

Private Sub ContinueButton_Click()
'Continueボタンが押下された時の処理
    'ここでは何もしない
    'CTestConditionControllerクラスでこのイベントを取得し、そこで処理する
    
End Sub

Private Sub ExecuteButton_Click()
'Executeボタンが押下された時の処理
    'ここでは何もしない
    'CTestConditionControllerクラスでこのイベントを取得し、そこで処理する

End Sub

Private Sub ReloadButton_Click()
'Reloadボタンが押下された時の処理
    'ここでは何もしない
    'CTestConditionControllerクラスでこのイベントを取得し、そこで処理する
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'×ボタンが押下された時の処理

    RaiseEvent QueryClose(Cancel, CloseMode)   '×ボタンでは終了できない旨をMsgBoxで表示する

End Sub

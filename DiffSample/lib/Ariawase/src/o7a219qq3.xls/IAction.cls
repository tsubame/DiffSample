VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   各アクションクラスが実装するインターフェイス
'
'目的:
'   アクションの実行やパラメータ等のログ出力を行う
'
'作成者:
'   0145206097
'
Option Explicit

Public Sub Execute()
End Sub

Public Property Get Name() As String
End Property

Public Property Get GroupIndex() As Long
End Property

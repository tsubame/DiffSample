VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IParameterViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   パラメーター表示用インターフェイス
'
'目的:
'   各目的ごとに必要となるビューワーが実装する
'
'作成者:
'   0145184306
'
Option Explicit

Public Sub View()
End Sub

Public Property Get EndStatus() As Boolean
End Property

Public Property Let EndStatus(ByVal pStatus As Boolean)
End Property

Public Property Get MyName() As String
End Property


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IActionStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   DC Test Scenario用のリーダー/ライターが実装するストリームインターフェイス
'
'目的:
'   DC Test Scenarioが各データをリード/ライトする際のデータアクセス制御を行う
'
'作成者:
'   0145206097

Option Explicit

Public Sub MoveNextCategory()
End Sub

Public Sub MoveNextGroup()
End Sub

Public Sub MoveNextData()
End Sub

Public Sub SetCategoryLocation(ByVal dataName As String)
End Sub

Public Sub SetGroupLocation(ByVal GroupIndex As Long)
End Sub

Public Sub SetDataLocation(ByVal dataIndex As Long)
End Sub

Public Function IsEndOfCategory() As Boolean
End Function

Public Function IsEndOfGroup() As Boolean
End Function

Public Function IsEndOfData() As Boolean
End Function

Public Sub Rewind()
End Sub

Public Sub SeekChangedLocation()
End Sub

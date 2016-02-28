VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFileStream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   リーダー/ライターが実装する汎用ストリームインターフェイス
'
'目的:
'   各データをリード/ライトする際のデータアクセス制御を行う
'
'作成者:
'   0145206097
'
Option Explicit

Public Sub MoveNext()
End Sub

Public Sub Rewind()
End Sub

Public Sub SetLocation(ByVal dataName As String)
End Sub

Public Function IsEOR() As Boolean
End Function

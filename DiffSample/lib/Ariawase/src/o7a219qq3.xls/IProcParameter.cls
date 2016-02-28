VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IProcParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   プロセッサが利用するインターフェース。
'
'目的:
'   各目的ごとに必要となるパラメータが実装する
'
'作成者:
'   0145184304
'
Option Explicit

Public Function Clone() As IProcParameter
End Function

Public Sub Dump(ByRef pWriter As IProcParamWriter)
End Sub

Public Property Get Comment() As String
End Property

Public Property Get Name() As String
End Property

Public Property Let Comment(ByVal pComment As String)
End Property

Public Property Let Name(ByVal TagName As String)
End Property

Public Sub Read(ByRef pReader As IProcParamReader)
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IParameterBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   パラメータバンククラスが実装するインターフェイス
'
'目的:
'
'作成者:
'   0145184304
'
Option Explicit

Public Sub Clear()
End Sub

Public Function Count() As Long
End Function

Public Sub Delete(ByVal paramName As String)
End Sub

Public Function Item(ByVal paramName As String) As IProcParameter
End Function

Public Function IsExist(ByVal paramName As String) As Boolean
End Function

Public Function Move(ByVal paramName As String) As IProcParameter
End Function

Public Sub Add(ByRef Param As IProcParameter, ByVal paramName As String, Optional ByVal pFixedPlane As Boolean = False)
End Sub

Public Function List() As String
End Function

Public Property Get IsOverwriteMode() As Boolean

End Property

Public Property Let IsOverwriteMode(ByRef mode As Boolean)

End Property
Public Function Keys() As Collection
End Function
Public Function FixedKeys() As Collection
End Function


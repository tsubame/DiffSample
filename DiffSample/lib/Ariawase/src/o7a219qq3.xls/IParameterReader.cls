VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IParameterReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   各リーダーが実装するインターフェイス
'
'目的:
'   各目的ごとに必要となるリーダーが実装する
'
'作成者:
'   0145206097
'
Option Explicit

Public Function ReadAsDouble(ByVal dataName As String) As Double
End Function

Public Function ReadAsString(ByVal dataName As String) As String
End Function

Public Function ReadAsInteger(ByVal dataName As String) As Integer
End Function

Public Function ReadAsLong(ByVal dataName As String) As Long
End Function

Public Function ReadAsBoolean(ByVal dataName As String) As Boolean
End Function

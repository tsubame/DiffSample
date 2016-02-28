VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IParameterWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   各ライターが実装するインターフェイス
'
'目的:
'   各目的ごとに必要となるライターが実装する
'
'作成者:
'   0145206097
'
Option Explicit

Public Sub WriteAsDouble(ByVal dataName As String, ByVal dblData As Double)
End Sub

Public Sub WriteAsString(ByVal dataName As String, ByVal strData As String)
End Sub

Public Sub WriteAsInteger(ByVal dataName As String, ByVal intData As Integer)
End Sub

Public Sub WriteAsLong(ByVal dataName As String, ByVal lngData As Long)
End Sub

Public Sub WriteAsBoolean(ByVal dataName As String, ByVal blnData As Boolean)
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   各パラメータクラスが実装するインターフェイス
'
'目的:
'
'作成者:
'   0145206097
'
Option Explicit

Public Sub Read(ByRef reader As IParameterReader)
End Sub

Public Sub WriteOut(ByRef writer As IParameterWriter)
End Sub

Public Property Let AsDouble(ByVal dblData As Double)
End Property

Public Property Get AsDouble() As Double
End Property

Public Property Let AsString(ByVal strData As String)
End Property

Public Property Get AsString() As String
End Property

Public Property Let AsLong(ByVal lngData As Long)
End Property

Public Property Get AsLong() As Long
End Property

Public Property Let AsBoolean(ByVal blnData As Boolean)
End Property

Public Property Get AsBoolean() As Boolean
End Property

Public Property Let LowerLimit(ByVal lLimit As Double)
End Property

Public Property Get LowerLimit() As Double
End Property

Public Property Let Name(ByVal dataName As String)
End Property

Public Property Get Name() As String
End Property

Public Property Let UpperLimit(ByVal uLimit As Double)
End Property

Public Property Get UpperLimit() As Double
End Property

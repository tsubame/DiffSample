VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IProcParamWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'概要:
'   各WriterUtilityが実装するインターフェイス
'
'目的:
'   各目的ごとに必要となるライターが実装する
'
'作成者:
'   0145184306
'

Public Sub WriteAsDouble(ByVal dataName As String, ByVal dblData As Double)
End Sub

Public Sub WriteAsInteger(ByVal dataName As String, ByVal intData As Integer)
End Sub

Public Sub WriteAsLong(ByVal dataName As String, ByVal lngData As Long)
End Sub

Public Sub WriteAsBoolean(ByVal dataName As String, ByVal blnData As Boolean)
End Sub

Public Sub WriteAsPlane(ByVal pParamName As String, ByRef pPlaneParam As CImgPlane)
End Sub

Public Sub WriteAsString(ByVal dataName As String, ByVal strData As String)
End Sub

Public Property Get MyName() As String
End Property

Public Sub WriteAsBlockEnd(ByVal pParamName As String)
End Sub

Public Sub WriteAsBlockStart(ByVal pParamName As String)
End Sub

Public Sub ReturnToRootBlock()
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDeviceParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   DeviceParameter取得するためのインターフェイス
'
'   Revision History:
'       Data        Description
'       2010/11/19  DeviceParameterを取得するI/F機能を実装した
'
'作成者:
'   0145184346
'

Option Explicit

Public Function AsAccTimeH() As Double()
End Function

Public Function AsAccTimeV() As Double()
End Function

Public Function AsDecibel() As Double()
End Function

Public Function AsDouble() As Double()
End Function

Public Function AsRatio() As Double()
End Function

Public Function AsString() As String()
End Function

Public Function AsStringAccTimeH() As String()
End Function

Public Function AsStringAccTimeV() As String()
End Function

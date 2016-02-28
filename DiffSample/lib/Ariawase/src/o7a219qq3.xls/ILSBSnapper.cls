VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILSBSnapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   換算用LSB値を取得するためのインターフェイス
'
'   Revision History:
'       Data        Description
'       2010/10/22  換算用LSB値を取得するI/F機能を実装した
'       2010/11/09  戻り値を変更した(CSiteArray→CDeviceParamArray)
'
'作成者:
'   0145184346
'

Option Explicit

Public Property Get Lsb() As CParamDoubleArray
End Property

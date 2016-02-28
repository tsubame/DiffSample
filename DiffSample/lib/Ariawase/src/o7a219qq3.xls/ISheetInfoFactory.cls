VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISheetInfoFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   CSheetInfoのFactoryのインターフェース
'
'目的:
'   CSheetInfoFactoryはこのインターフェースを実装しなければならない
'
'作成者:
'   a_oshima

Option Explicit

Public Function CreateSheetInfo(ByRef pSheet As Worksheet) As CSheetInfo
End Function






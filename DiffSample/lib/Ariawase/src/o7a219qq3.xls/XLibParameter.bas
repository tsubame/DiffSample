Attribute VB_Name = "XLibParameter"
'概要:
'   パラメータクラス用オブジェクト作成モジュール
'
'目的:
'   パラメータオクラスブジェクトを作成し返す（ライブラリ作成ルールに基づく）
'
'作成者:
'   0145206097
'
Option Explicit







Public Function CreateCParamDouble() As CParamDouble
    Set CreateCParamDouble = New CParamDouble
End Function

Public Function CreateCParamLong() As CParamLong
    Set CreateCParamLong = New CParamLong
End Function

Public Function CreateCParamBoolean() As CParamBoolean
    Set CreateCParamBoolean = New CParamBoolean
End Function

Public Function CreateCParamName() As CParamName
    Set CreateCParamName = New CParamName
End Function

Public Function CreateCParamString() As CParamString
    Set CreateCParamString = New CParamString
End Function

Public Function CreateCParamStringWithUnit() As CParamStringWithUnit
    Set CreateCParamStringWithUnit = New CParamStringWithUnit
End Function

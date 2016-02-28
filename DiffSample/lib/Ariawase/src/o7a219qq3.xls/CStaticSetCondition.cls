VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStaticSetCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   静的条件設定クラス
'
'目的:
'   TestConditionシートから取得したConditionを従来の処理で実行するクラス
'
'作成者:
'   0145184935 K.Hara
'

Option Explicit

Implements ISetCondition

Private Sub Class_Initialize()
'
End Sub

Private Sub Class_Terminate()
'
End Sub

Public Sub IsetCondition_Execute(ByVal pFunctionInfo As CAcquireFunctionInfo)
'内容:
'   従来の処理でマクロ実行を行う
'
'[pFunctionInfo] IN CAcquireFunctionInfo型: FWSetCondition行
'
'備考:
'
    '従来の処理でマクロ実行を行う
    Call TheCondition.SetCondition(pFunctionInfo.InstanceName)

End Sub

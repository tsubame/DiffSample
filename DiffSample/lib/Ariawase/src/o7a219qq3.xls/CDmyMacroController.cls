VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDmyMacroController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   マクロ無効設定時に、測定条件設定用のマクロを実行しない為の偽物機能
'
'目的:
'   マクロ有効無効切り替え対応
'

Option Explicit

Implements IMacroController 'マクロ実行機能用I/F

'I/F用の実装
Private Sub IMacroController_RunMacro(ByRef pParameters As CSetFunctionInfo)
'内容:
'   Excelマクロの実行処理
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
End Sub

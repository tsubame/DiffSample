VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExcelMacroController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   測定条件設定用のマクロを実行する
'
'目的:
'   マクロ有効無効切り替え対応
'

Option Explicit

Implements IMacroController                           'マクロ実行機能用I/F
Private m_TargetBookName As String                    'マクロの存在するワークブック名

Public Sub Initialize()
'内容:
'   全体機能の初期化
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    
    m_TargetBookName = ""

End Sub

Public Property Let TargetBookName(ByVal pBookName As String)
'内容:
'   マクロの存在するWorkBook名
'
'パラメータ:
'   [pBookName]  In  String型:   ワークブック名
'
'戻り値:
'
'注意事項:
'
    
    m_TargetBookName = pBookName

End Property

Public Property Get TargetBookName() As String
'内容:
'   マクロの存在するWorkBook名
'
'パラメータ:
'
'戻り値:
'   ワークブック名
'
'注意事項:
'
    
    TargetBookName = m_TargetBookName

End Property

'I/F用の実装
Private Sub IMacroController_RunMacro(ByRef pParameters As CSetFunctionInfo)
'内容:
'   Excelマクロの実行
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    
    Call Excel.Application.Run("'" & m_TargetBookName & "'!" & pParameters.FunctionName, pParameters)

End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMacroExecInterceptor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   マクロ実行時にログを取得するためのラッパークラス
'
'目的:
'   ログを取得しながらマクロ実行を行う
'

Option Explicit

Implements IMacroController 'マクロ実行機能用I/F

Private m_Logger As CSetConditionLogger        'ロガーObject
Private m_MacroController As IMacroController  'マクロ実行機能提供Object

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
    
    Set m_Logger = Nothing
    Set m_MacroController = Nothing

End Sub

Public Property Set SetConditionLogger(ByRef pLoggerObject As CSetConditionLogger)
'内容:
'   条件設定履歴記録用ロガーObject
'
'パラメータ:
'   [pLoggerObject]  In  CSetConditionLogger型: ロガーObject

'戻り値:
'
'注意事項:
'
    
    Set m_Logger = pLoggerObject

End Property

Public Property Get SetConditionLogger() As CSetConditionLogger
'内容:
'   条件設定履歴記録用ロガーObject
'
'パラメータ:

'戻り値:
'   ロガーObject
'
'注意事項:
'
    
    Set SetConditionLogger = m_Logger

End Property

Public Property Get MacroController() As IMacroController
'内容:
'   マクロ実行機能Object
'
'パラメータ:

'戻り値:
'   マクロ実行機能提供Object
'
'注意事項:
'
    
    Set MacroController = m_MacroController

End Property

Public Property Set MacroController(ByRef pController As IMacroController)
'内容:
'   マクロ実行機能Object
'
'パラメータ:
'   [pController]  In  IMacroController型: マクロ実行機能提供Object
'
'戻り値:
'
'注意事項:
'
    
    Set m_MacroController = pController

End Property

Private Sub IMacroController_RunMacro(ByRef pParameters As CSetFunctionInfo)
'内容:
'   マクロの実行処理
'
'パラメータ:
'   [pParameters]  In  CSetFunctionInfo型: マクロ実行に使用するパラメータ
'
'戻り値:
'
'注意事項:
'
    
    Dim ExecutionTime As Double
    ExecutionTime = 0#
    
    '実行時間を計測しながらマクロを実行する
    If pParameters.IsEnableMacro = True Then
        Dim TimerObj As CEeeTimer
        Set TimerObj = New CEeeTimer
        With TimerObj
            Call .Initialize
            Call .StartTimer
            'Excelマクロの実行
            Call m_MacroController.RunMacro(pParameters)
            ExecutionTime = .StopTimer
        End With
    End If

    '実行ログ記録処理
    Call m_Logger.AddRecord(ExecutionTime, pParameters)
    
End Sub

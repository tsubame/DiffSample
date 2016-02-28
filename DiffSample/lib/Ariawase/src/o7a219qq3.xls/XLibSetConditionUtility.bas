Attribute VB_Name = "XLibSetConditionUtility"
'概要:
'   測定条件設定機能提供
'
'目的:
'   ワークシートのデータを使用した測定条件設定の実現
'   測定条件設定機能のイニシャライズ関係の記述
'
'作成者:
'   SLSI 今手
'   tomoyoshi.takase
'
'注意点:
'   初回の利用時に、Initializeの実行が実用です。
'   TheErrorとして公開されているエラーマネージャObjectが必要です
'   変更履歴
'　 2010/03/08、アドイン版で作成されていたものを、モジュールで動作するように変更。
'　 AddWorkSheetメソッドはアドインでなくなったので廃止
'

Option Explicit

'エラー時の情報定義
Private Const ERR_NUMBER = 9999                              'エラー時に渡すエラー番号
Private Const CLASS_NAME = "XLibSetConditionUtilities" 'このクラスの名前

'公開機能 Object
Public TheCondition As CTestConditionManager

Private mSaveFileName As String

'公開Eee-JOBワークシートの定義用
Enum EEE_CONDITION_WORKSHEET
    TestCondition_EeeJobSheet = 0
End Enum

Public Sub CreateTheConditionIfNothing()
'内容:
'   As Newの代替として初回に呼ばせるインスタンス生成処理
'
'パラメータ:
'   なし
'
'注意事項:
'
    On Error GoTo ErrHandler
    If TheCondition Is Nothing Then
        '######## TestCondition Block ########
        Call CreateTestCondition(ThisWorkbook)
        '測定条件表の初期設定
        TheCondition.TestConditionSheet = GetWkShtReaderManagerInstance.GetActiveSheetName(eSheetType.shtTypeTestCond)        '条件表ワークシート名を設定する
        Call TheCondition.LoadCondition
    End If
    Exit Sub
ErrHandler:
    Set TheCondition = Nothing
    TheError.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub


Public Sub SetLogModeTheCondition(ByVal pEnableLoggingTheCondition As Boolean, Optional saveFileName As String = "EeeJOBLogSetCondition.csv")
    Call TheCondition.LoadCondition
    TheCondition.CanHistoryRecord = pEnableLoggingTheCondition
    mSaveFileName = saveFileName
End Sub


'ライブラリの初期化
Public Sub CreateTestCondition(ByVal pJobWorkBook As Workbook)
'内容:
'   EeeTestCondition全体の初期化
'
'パラメータ:
'   [pTheErrorObj]  In  Object型:     エラー管理機能提供Object
'   [pJobWorkBook]  In  Workbook型:   JOBのWorkbook
'
'戻り値:
'
'注意事項:
'
    '測定条件マネージャの初期化
    Call InitTestConditionManager(pJobWorkBook.Name)
    
End Sub

Public Sub DestroyTestCondition()
'内容:
'   EeeTestConditionAddIn全体の終了処理
'
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    Set TheCondition = Nothing

End Sub

Public Function GetTestConditionInstance() As CTestConditionManager
'内容:
'   測定条件設定機能Objectの提供
'
'パラメータ:
'
'戻り値:
'   測定条件設定機能Object
'
'注意事項:
'   Initializeが未実行の状態で、本命令を実行するとエラーとなります。
'
'
    If TheCondition Is Nothing Then
        Call TheError.Raise(9999, CLASS_NAME, "Initialization is unexecution." & " @EeeTestConditionAddIn")
        'Call CreateTestCondition(ThisWorkbook)
    Else
        Set GetTestConditionInstance = TheCondition
    End If
End Function

Public Function RunAtJobEnd() As Long
    If Not (TheCondition Is Nothing) Then
        If TheCondition.CanHistoryRecord Then
            Call TheCondition.SaveHistoryLog(mSaveFileName)
            Call TheCondition.ClearExecHistory
        End If
        TheCondition.CanHistoryRecord = False
    End If
End Function


'------------------------------------------------------------------------------------------------
'以下 Private

'インスタンスの生成と初期化の処理
Private Sub InitTestConditionManager(ByVal pJobWorkbookName As String)
'内容:
'
'作成者:
'  tomoyoshi.takase
'パラメータ:
'   [pJobWorkbookName]  In  1):対象ワークブック名
'戻り値:
'
'注意事項:
'

    Set TheCondition = New CTestConditionManager
    Call TheCondition.Initialize
    TheCondition.JobWorkbookName = pJobWorkbookName
End Sub

Public Sub ChangeDefaultSettingTheCondition()
'内容:
'   選択しているグループをDefaultに変更するように要請する
'パラメータ:
'
'戻り値:
'
'注意事項:
'
    Call TheCondition.ChangeDefaultSetting

End Sub


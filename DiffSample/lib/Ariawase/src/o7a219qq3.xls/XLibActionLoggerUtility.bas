Attribute VB_Name = "XLibActionLoggerUtility"
'概要:
'   ActionLoggerのユーティリティ
'
'目的:
'   ActionLoggerの初期化/破棄のユーテリティを定義する
'
'作成者:
'   a_oshima

Option Explicit

Private mEnableActionLogger As Boolean
Private mActionLogger As CActionLogger   ' 実行LogデータをダンプするActionLoggerを保持する
Private mActionLogFileName As String
Private Const DEFAULT_LOGNAME As String = "EeeJOBLogActionLogger.csv"

Public Sub CreateActionLoggerIfNothing()
'内容:
'   As Newの代替として初回に呼ばせるインスタンス生成処理
'   履歴のクリアも行う
'
'パラメータ:
'   なし
'
'注意事項:
'
    On Error GoTo ErrHandler
    If mActionLogger Is Nothing Then
        Set mActionLogger = New CActionLogger
    End If
    Call mActionLogger.Initialize
    Exit Sub
ErrHandler:
    Set mActionLogger = Nothing
    TheError.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Sub

Public Sub SetLogModeActionLogger(ByVal pEnableLoggingActionLogger As Boolean, Optional ByVal pSaveActionLogFileName As String = DEFAULT_LOGNAME)

    If TheExec.JobIsValid = False Then
        MsgBox "The log mode cannot be set before Validation."
        Exit Sub
    End If

    Dim YesOrNo As VbMsgBoxResult
    If TheExec.RunMode = runModeProduction Then
        If pEnableLoggingActionLogger = True Then
            YesOrNo = MsgBox("Now ProductionMode! But SetLogModeActionLogger is True!" & vbCrLf & _
                            "EeeJOB will output DataLog." & vbCrLf & _
                            "Please make sure." & vbCrLf & _
                            "When You don't want to output This MsgBox, Please change runModeProduction => runModeDebug" & vbCrLf & _
                            "" & vbCrLf & _
                            "Output DataLog?" & vbCrLf _
                            , vbYesNo + vbQuestion, "Confirm Output DataLog(SetLogModeActionLogger)")
            If YesOrNo = vbYes Then
                mEnableActionLogger = True
            Else
                mEnableActionLogger = False
            End If
        Else
            mEnableActionLogger = False
        End If
    Else
        mEnableActionLogger = pEnableLoggingActionLogger
    End If

    If mEnableActionLogger = True Then
        TheExec.Datalog.WriteComment "Eee JOB Output Log! :Action Logger"
    End If

    mActionLogFileName = pSaveActionLogFileName

    If TheExec.Flow.IsRunning = True Then
        ApplyLogModeActionLogger
    End If

End Sub

Sub ApplyLogModeActionLogger()

    If mEnableActionLogger = True And mActionLogger Is Nothing Then
            CreateActionLoggerIfNothing
            Call XLibTOPT_FW.EnableInterceptor(mEnableActionLogger, mActionLogger)
            Call XLibImageEngineUtility.EnableInterceptor(mEnableActionLogger, mActionLogger)
#If ITS <> 0 Then
            Call EnableScenarioActionLogger(mEnableActionLogger, mActionLogger)
#End If
    ElseIf mEnableActionLogger = False And Not mActionLogger Is Nothing Then
            Call XLibTOPT_FW.EnableInterceptor(mEnableActionLogger, mActionLogger)
            Call XLibImageEngineUtility.EnableInterceptor(mEnableActionLogger, mActionLogger)
#If ITS <> 0 Then
            Call EnableScenarioActionLogger(mEnableActionLogger, mActionLogger)
#End If
            DestroyActionLogger
    End If
End Sub

Public Sub DestroyActionLogger()
    Set mActionLogger = Nothing
End Sub

Public Function GetActionLoggerInstance() As CActionLogger
'内容:
'   ActionLoggerのインスタンスを返す
'
'パラメータ:
'   なし
'
'戻り値:
'   ActionLoggerのインスタンス
'
'例外:
'   未初期化時に呼ばれるとVBA例外発生
'  （パフォーマンス改善のためAsNewの代替として用意してあり、Nothingチェックは行わない）
'
'注意事項:
'   初期化処理を先に呼び、インスタンスが生成されていること

    'CreateActionLoggerIfNothing()
    Set GetActionLoggerInstance = mActionLogger
End Function

Public Function RunAtJobEnd() As Long
    If Not mActionLogger Is Nothing Then
        If mEnableActionLogger = True Then
            SaveActionLoggerHistoryLog
            Call XLibTOPT_FW.EnableInterceptor(False, mActionLogger)
            Call XLibImageEngineUtility.EnableInterceptor(False, mActionLogger)
#If ITS <> 0 Then
            Call EnableScenarioActionLogger(False, mActionLogger)
#End If
            DestroyActionLogger
            mEnableActionLogger = False
        End If
    End If
End Function



'ImageEngineからインポート
Public Function GetActionLoggerEnable() As Boolean
'とりあえず、常にTrue
'    GetActionLoggerEnable = mActionLogger.EnableLogging
    GetActionLoggerEnable = mEnableActionLogger
End Function

Public Function GetActionLogFileName() As String
    GetActionLogFileName = mActionLogFileName
End Function

Public Property Let SetActionLogFileName(ByRef strName As String)
    mActionLogFileName = strName
End Property

Public Function SaveActionLoggerHistoryLog()
    If mActionLogFileName <> "" Then
        Call mActionLogger.SaveHistoryLog(mActionLogFileName)
    Else
        Call mActionLogger.SaveHistoryLog(makeFileName)
    End If
End Function

Public Function ClearActionLoggerHistoryLog()
    Call mActionLogger.ClearHistories
End Function

Private Function makeFileName() As String
    makeFileName = getCurrentDirectory & "\" & getToday & "_" & "TOPT_ActionLog" & ".csv"
End Function

Private Function getToday() As String
    getToday = Format$(DateTime.Now, "yyyymmdd") & "_" & Format$(DateTime.Now, "hhnnss")
End Function

Private Function getCurrentDirectory() As String
    getCurrentDirectory = ActiveWorkbook.Path
End Function



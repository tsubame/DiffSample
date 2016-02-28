Attribute VB_Name = "XLibToptFrameWorkUtility"
'概要:
'   ToptFrameWorkのユーティリティ
'
'目的:
'   XLibIGXLEventsのOn～メソッドで呼ばせたい関数を定義する
'
'作成者:
'   a_oshima

Option Explicit

Private mFailedJobInitialize As Boolean

Public Function ResetEeeJobObjects() As Long
'概要:
'   OnProgramLoadedで呼ばせたい関数
'
    On Error GoTo ErrHandler
    Call DestroyAllEeeJobObjects

    Exit Function
ErrHandler:
    MsgBox "Failed at Eee-Job Initialize Routine: " & CStr(Err.Number) & " - " & Err.Source & vbCrLf & vbCrLf & Err.Description
End Function

Public Function ResetEeeJobSheetObjects() As Long
'概要:
'   OnProgramValidatedで呼ばせたい関数
'

    On Error GoTo ErrHandler
    Call DestroyAllEeeJobSheetObjects

    Exit Function
ErrHandler:
    MsgBox "Failed at Eee-Job Initialize Routine: " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
End Function

Public Function RunAtValidated() As Long
'概要:
'   OnProgramValidatedで呼ばせたい関数
'

    On Error GoTo ErrHandler

    Call CreateTheErrorIfNothing
    Call CreateReaderManagerIfNothing
    Call CreateTheParameterBankIfNothing        'add 20110209
    Call CreateTheIDPIfNothing
    Call CreatePlaneMapIfNothing
    Call CreateKernelManagerIfNothing
    Call CreateTheImageTestIfNothing
    Call CreateTheConditionIfNothing
    Call XLibSetConditionUtility.ChangeDefaultSettingTheCondition
#If ITS <> 0 Then
    Call XLibImpUIControllerUtility.RunAtValidated
#End If
    Exit Function
ErrHandler:
    MsgBox "Failed at Eee-Job Initialize Routine: " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
End Function

Public Function RunAtJobStart() As Long
'概要:
'   OnProgramStartedで呼ばせたい関数
'
    mFailedJobInitialize = False

    On Error GoTo ErrHandler
    '### インスタンスが無ければ生成 #######################
    Call CreateTheErrorIfNothing
    Call CreateReaderManagerIfNothing
    Call CreateTheParameterBankIfNothing        'add 20110209
    Call CreateTheSystemInfoIfNothing
    Call CreateTheIDPIfNothing
    Call CreateTheVarBankIfNothing
    Call CreateTheFlagBankIfNothing

    Call CreatePlaneMapIfNothing
    Call CreatePMDIfNothing
    Call CreateKernelManagerIfNothing
    Call CreateTheImageTestIfNothing
#If ITS <> 0 Then
    Call CreateScenarioBuilderIfNothing
    Call CreateTheImgTestScenarioIfNothing
#End If
    Call CreateTheConditionIfNothing
    Call CreateTOPTFWIfNothing
    Call CreateTheDeviceProfilerIfNothing
    Call TheParameterBank.Clear
    Call InitTestScenario

    '### テスト毎のリセット作業 ###########################
    TheIDP.ResetTest
    XLibTOPT_FW.ResetStatus

    '### 各ロガーをDisableに設定 ##########################
    'TheVarBankのログ機能は開発者向けとして封印
'    Call SetLogModeActionLogger(True)
'    Call SetLogModeTheIDP(True)
'    Call SetLogModeTheCondition(True)

'    Call SaveModeTheVarBank(True)

    XLibActionLoggerUtility.ApplyLogModeActionLogger
    
#If ITS <> 0 Then
    Call XLibImpUIControllerUtility.RunAtJobStart
    Call PrepareDumpDirectory(ThisWorkbook.Path & Application.PathSeparator & "Dump")
#End If

    Exit Function
ErrHandler:
    MsgBox "Failed at Job Initialize Routine: " & CStr(Err.Number) & " - " & Err.Source & Chr$(13) & Chr$(13) & Err.Description
    Call AbortJob
End Function

Public Function RunAtJobEnd() As Long
'概要:
'   OnProgramEndedで呼ばせたい関数

'{ ログレポートライターの後処理
    Call CloseDcLogReportWriter
'}
    Call XLibDcScenarioLoopOption.RunAtJobEnd
    
    Call XLibActionLoggerUtility.RunAtJobEnd
    Call XLibImageEngineUtility.RunAtJobEnd
#If ITS <> 0 Then
    Call XLibImSceEngineUtility.RunAtJobEnd
    Call XLibScenarioUtility.RunAtJobEnd
#End If
    Call XLibTheVarBankUtility.RunAtJobEnd
    Call XLibImgUtility.RunAtJobEnd
    Call XLibSetConditionUtility.RunAtJobEnd
    Call XLibTheFlagBankUtility.RunAtJobEnd
    Call XLibTheDeviceProfilerUtility.RunAtJobEnd
    Call XLibTheParameterBankUtility.RunAtJobEnd

    Unload ScenarioParameterViewer
#If ITS <> 0 Then
    Call XLibImpUIControllerUtility.RunAtJobEnd
#End If

'    TheIDP.ResetTest
    '### TheErrorのログ機能は開発者向けとして封印 #########
'    Call XLibErrManangerUtility.RunAtJObEnd(pSaveFileName:="TheErrorLog.csv")

End Function

Public Sub AbortJob()
    mFailedJobInitialize = True
End Sub

Public Function FailedJobInitialize() As Boolean
    FailedJobInitialize = mFailedJobInitialize
End Function

Public Sub DestroyAllEeeJobObjects()
    Call DestroyTheError
    Call DestroyTheIDP
    Call DestroyPMDSheet
    Call DestroyActionLogger
#If ITS <> 0 Then
    Call DestroyImpUIController
#End If

    Call DestroyAllEeeJobSheetObjects
End Sub

Public Sub DestroyAllEeeJobSheetObjects()
    Call DestroyTheVarBank
    Call DestroyTheFlagBank
    Call DestroyWkShtReaderManager
#If ITS <> 0 Then
    Call DestroyTheImgTestScenario
    Call DestroyScenarioBuilder
#End If
    Call DestroyTheImageTest
    Call DestroyTestCondition
    Call DestroyTOPTFW
    Call DestroyTheDeviceProfiler
End Sub


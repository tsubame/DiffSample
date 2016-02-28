Attribute VB_Name = "XEeeAuto_TestConditionMacro"
'概要:
'   TestConditionから呼ばれるマクロ集
'
'目的:
'   TestConditionシートから呼ばれるマクロを定義する
'
'作成者:
'   2011/12/07 Ver0.1 D.Maruyama    Draft
'   2011/12/15 Ver0.2 D.Maruyama    FW_CallUserMacro追加
'   2011/12/16 Ver0.3 D.Maruyama    以下の関数にOptional引数としてWaitを追加
'                                   ・FW_set_voltage
'                                   ・FW_PatRun
'                                   ・FW_PatSet
'   2012/01/23 Ver0.4 D.Maruyama    SUB電流測定時のパラメータ設定関数を追加
'                                   ・FW_SetSubCurrentParam
'                                   ・GetSubCurrentAverageCount（取得用）
'                                   ・GetSubCurrentClampCurrent（取得用）
'                                   ・GetSubCurrentWaitTime（取得用）
'                                   画像キャプチャのパラメータ設定関数を追加
'                                   ・FW_SetCaptureAverage
'                                   ・GetCaptureAverageCount（取得用）
'                                   ・GetCaptureAverageMode（取得用）
'   2012/02/03 Ver0.5 D.Maruyama    SUB電流測定時のKey名をTestInstaceから取得するように変更
'                                   画像キャプチャのパラメータ設定関数にFrameSkip数に追加し、関数をリネーム
'                                   ・FW_SetCaptureParam
'                                   ・GetCaptureParamAverageCount（取得用）
'                                   ・GetCaptureParamAverageMode（取得用）
'                                   FrameSkip取得用の関数を追加
'                                   ・GetCaptureParamFrameSkip（取得用）
'   2012/02/09 Ver0.6 D.Maruyama    FW_set_voltageについて以下の対処を行い変数が2つ減る
'                                   ・SUB切り離し業務をJobRouteに移管
'                                   ・XCLRの処理をTestConditionに移管
'   2012/02/09 Ver0.7 D.Maruyama    以下の関数はインスタンス名をキー名に付与した形に変更（複数SUB測定に対応するため）
'                                   ・FW_SetSubCurrentParam
'                                   ・GetSubCurrentAverageCount（取得用）
'                                   ・GetSubCurrentClampCurrent（取得用）
'                                   ・GetSubCurrentWaitTime（取得用）
'   2012/03/06 Ver0.71 D.Maruyama    エラーハンドリングをすべてOnError形式に統一
'                                   以下の関数を削除
'                                   ・SiteLimit
'   2012/03/07 Ver0.8 D.Maruyama    以下の関数を追加
'                                   ・FW_WaitSetTopt
'                                     新規作成､従来のFW_WaitSet｡
'                                   ・GetScrnMeasureWaitParam
'                                   ・FW_ScrnMeasureWaitParam
'
'                                   以下の関数を修正
'                                   ・FW_SET_RELAY_CONDITION
'                                     CUBの引数をとるように変更
'                                   ・FW_DisconnectPins
'                                     複数セルにまたいで記述できるように変更
'                                   ・FW_ConnectPins
'                                     FW_ConnectIOPinsに変更｡PinsConnectのみ
'                                   ・FW_WaitSet
'                                     まちかたをTheHdw.Waitに固定
'                                   ・FW_SetFVMI_BPMU
'                                   ・FW_SetFIMV_BPMU
'                                   ・FW_DisconnectPins_BPMU
'                                     サイト指定を省略可能にした｡省略時は全サイト実行
'                                   ・FW_PatternStop
'                                     名称をFW_StopPatに変更
'                                   ・ConvertStartState
'                                     chStartFmt , chStartNoneを追加
'   2012/03/23 Ver0.9 D.Maruyama    以下の関数を削除
'                                   ・GetScrnMeasureWaitParam
'                                   ・FW_ScrnMeasureWaitParam
'                                   以下の関数を追加
'                                   ・FW_SetScrnMeasureParam
'                                   ・GetScrnMeasureWaitTime
'                                   ・GetScrnMeasureAverageCount
'                                   ・FW_SeparateFailSiteGnd
'   2012/04/06 Ver1.0 D.Maruyama    以下の関数を修正
'                                   ・FW_SeparateFailSiteGnd
'                                       SUBの取得を関数の先頭に移動、省略文字の場合すぐに抜けるようにした
'                                   ・FW_set_voltage
'                                       PowerConditionシート廃止に伴い、引数を、条件名と、シーケンス名の2つの引数に分割
'   2012/04/09 Ver1.1 D.Maruyama    FW_SeparateFailSiteGnd関数でFailSiteはGNDリレーも切り離すように追加
'   2012/09/27 Ver1.2 H.Arikawa     以下の関数を追加
'                                   ・FW_SetGND
'   2012/10/01 Ver1.3 H.Arikawa     以下の関数のStop処理を削除
'                                   ・FW_SET_RELAY_CONDITION
'                                   ・FW_OptSet
'                                   ・FW_set_voltage
'                                   ・FW_ConnectIOPins
'                                   ・FW_DisconnectPins
'                                   ・FW_WaitSet
'                                   ・FW_WaitSetTopt
'                                   ・FW_SetFVMI
'                                   ・FW_SetFIMV
'                                   ・FW_SetFVMI_BPMU
'                                   ・FW_SetFIMV_BPMU
'                                   ・FW_DisconnectPins_BPMU
'                                   ・FW_PatSet
'                                   ・FW_PatRun
'                                   ・FW_StopPat
'                                   ・FW_SetIOPinState
'                                   ・FW_SetIOPinElectronics
'                                   ・FW_CallUserMacro
'                                   ・FW_SeparateFailSiteGnd
'                                   ・FW_SetSubCurrentParam
'                                   ・GetSubCurrentAverageCount
'                                   ・GetSubCurrentClampCurrent
'                                   ・GetSubCurrentWaitTime
'                                   ・FW_SetScrnMeasureParam
'                                   以下の関数のStop処理を削除し、テスト停止を追加
'                                   FW_SetCaptureParam
'                                   ・GetCaptureParamAverageCount
'                                   ・GetCaptureParamAverageMode
'                                   ・GetCaptureParamFrameSkip
'                                   ・GetScrnMeasureWaitTime
'                                   ・GetScrnMeasureAverageCount
'                                   ・FW_SetGND
'   2012/10/18 Ver1.4 H.Arikawa     以下の関数を追加
'                                   ・FW_PatRun_Decoder
'                                   ・FW_PatSet_Decoder
'   2012/10/19 Ver1.5 K.Tokuyoshi   以下の関数を追加
'                                   ・GetSubCurrentPinResourceName
'                                   ・FW_SET_RELAY_ON
'                                   ・FW_SET_RELAY_OFF
'                                   ・DutConnectDbNumber
'   2012/10/22 Ver1.8 K.Tokuyoshi   以下の関数を追加
'                                   ・FW_SetHoldVoltageParam
'                                   ・GetHoldVoltageAverageCount
'                                   ・GetHoldVoltageClampCurrent
'                                   ・GetHoldVoltageWaitTime
'   2012/10/22 Ver1.9 K.Tokuyoshi   以下の関数を修正
'                                   ・FW_SetSubCurrentParam
'   2012/11/02 Ver2.0 H.Arikawa     以下の関数を追加
'                                   ・FW_set_voltage_ForUS
'                                   ・PowerDownAndDisconnect
'                                   以下の関数を修正
'                                   ・FW_set_voltage
'   2012/12/20 Ver2.1 H.Arikawa     以下の関数を追加
'                                   ・FW_PatStatus
'   2013/01/14 Ver2.2 H.Arikawa     以下の関数を修正
'                                   ・FW_PatStatus
'                                   ・FW_PatSetTypeSelect
'   2013/01/22 Ver2.3 H.Arikawa     以下の関数を修正
'                                   ・FW_SetCaptureParam
'   2013/01/29 Ver2.4 H.Arikawa     以下の関数を暫定追加
'                                   ・FW_OptEscape
'                                   ・FW_OptModOrModZ1
'                                   ・FW_OptModOrModZ2
'   2013/01/31 Ver2.5 H.Arikawa     以下の関数を削除(未使用の為)
'                                   ・FW_SeparateFailSiteGnd
'   2013/02/01 Ver2.6 H.Arikawa     以下の関数を修正・追加
'                                   ・FW_SetSubCurrentParam
'                                   ・FW_SetSubCurrentParam_BPMU
'   2013/02/07 Ver2.7 H.Arikawa     以下の関数を修正
'                                   ・FW_SetCaptureParam
'   2013/02/08 Ver2.8 H.Arikawa     以下の関数を修正
'                                   ・FW_OptEscape
'                                   ・FW_OptModOrModZ1
'                                   ・FW_OptModOrModZ2
'   2013/02/12 Ver2.9 H.Arikawa     以下の関数を修正
'                                   ・FW_SetFIMV
'   2013/02/19 Ver3.0 H.Arikawa     以下の関数を修正(暫定)
'                                   ・FW_OptEscape
'                                   ・FW_OptModOrModZ1
'                                   ・FW_OptModOrModZ2
'   2013/02/22 Ver3.1 H.Arikawa     以下の関数を修正
'                                   ・FW_ScreeningWait
'   2013/02/22 Ver3.2 H.Arikawa     以下の関数にFlg_Simulatorの処理を追加
'                                   ・FW_PatSet
'                                   ・FW_PatSet_Decoder
'                                   ・FW_PatRun
'                                   ・FW_PatRun_Decoder
'                                   ・FW_StopPat
'                                   ・FW_PatSetTypeSelect
'                                   ・FW_PatStatus
'   2013/02/25 Ver3.3 H.Arikawa     以下の関数の確定版を追加
'                                   ・FW_OptEscape
'                                   ・FW_OptModOrModZ1
'                                   ・FW_OptModOrModZ2
'   2013/02/28 Ver3.4 H.Arikawa     以下の関数を修正
'                                   ・FW_PatSet
'   2013/03/04 Ver3.5 H.Arikawa     以下の関数を修正
'                                   ・FW_set_voltage
'                                   以下の関数を削除
'                                   ・FW_set_voltage_ForUS
'   2013/03/04 Ver3.6 H.Arikawa     以下の関数を修正
'                                   ・FW_set_voltage
'   2013/03/11 Ver3.7 H.Arikawa     以下の関数を簡略化
'                                   ・FW_OptEscape
'                                   ・FW_OptModOrModZ1
'                                   ・FW_OptModOrModZ2
'   2013/03/11 Ver3.8 H.Arikawa     以下の関数を追加
'                                   ・FW_DebugWait
'   2013/03/15 Ver3.9 H.Arikawa     以下の関数を追加
'                                   ・PatSet
'   2013/08/21 Ver4.0 H.Arikawa     以下の関数を追加
'                                   ・FW_PatSetCustomMacroA
'   2013/08/23 Ver4.1 H.Arikawa     以下の関数の光源SKIP FLAG ON時のSkip処理追加
'                                   ・FW_OptEscape
'                                   ・FW_OptModOrModZ1
'                                   ・FW_OptModOrModZ2
'   2013/09/27 Ver4.2 H.Arikawa     FW_SetSubCurrentParamとFW_SetSubCurrentParam_BPMUを統合
'                                   ・FW_SetSubCurrentParam
'   2013/10/28 Ver4.3 H.Arikawa     条件設定省略のフラグ化
'   2013/11/05 Ver4.4 T.Morimoto    FW_DcTopt_SetとFW_DcTopt_Measureを追加


Option Explicit

'VarBankやり取り用
Private Const SUBCURRENT_AVERAGE_COUNT As String = "_SUBCURRENT_AVERAGE_COUNT__"
Private Const SUBCURRENT_CLAMP_CURRENT As String = "_SUBCURRENT_CLAMP_CURRENT__"
Private Const SUBCURRENT_WAIT_TIME As String = "_SUBCURRENT_WAIT_TIME__"
Private Const SUBCURRENT_PIN_RESOURCE As String = "_SUBCURRENT_PIN_RESOURCE__"

Private Const HOLDVOLTAGE_AVERAGE_COUNT As String = "_HOLDVOLTAGE_AVERAGE_COUNT__"
Private Const HOLDVOLTAGE_CLAMP_CURRENT As String = "_HOLDVOLTAGE_CLAMP_CURRENT__"
Private Const HOLDVOLTAGE_WAIT_TIME As String = "_HOLDVOLTAGE_WAIT_TIME__"

Private Const SCRN_MEAS_WAIT_TIME As String = "_SCRNMEAS_WAIT_TIME__"
Private Const SCRN_MEAS_AVERAGE_COUNT As String = "_SCRNMEAS_AVERAGE_COUNT__"

Private Const CAPTURE_PARAM_AVERAGE_COUNT As String = "_CAPTURE_PARAM_AVERAGE_COUNT__"
Private Const CAPTURE_PARAM_AVERAGE_MODE As String = "_CAPTURE_PARAM_AVERAGE_MODE__"
Private Const CAPTURE_PARAM_FRAME_SKIP As String = "_CAPTURE_PARAM_FRAME_SKIP__"

Private Const CAPTURE_AVERAGE_MODE_AVERAGE As String = "Average"
Private Const CAPTURE_AVERAGE_MODE_NO_AVERAGE As String = "NonAverage"

Private OptCheckCounter As Double
Public PatCheckCounter As Double

'内容:
'   リレー設定を行う
'
'パラメータ:
'    [Arg0]      In   リレーセット名
'
'戻り値:
'
'注意事項:
'
Public Sub FW_SET_RELAY_CONDITION(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler

    If Parameter.ArgParameterCount() <> 3 Then
        Err.Raise 9999, "FW_SET_RELAY_CONDITION", "The number of FW_SET_RELAY_CONDITION's arguments is invalid." & " @ " & Parameter.ConditionName
    End If

'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_APMU_UB
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================

    'リレー設定
    Call SET_RELAY_CONDITION(Parameter.Arg(0), Parameter.Arg(1))
    
    If Parameter.Arg(2) <> "-" And Parameter.Arg(2) <> "" Then
        Application.Run Parameter.Arg(2)
    End If
    
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================

    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   光源設定を行う
'
'パラメータ:
'    [Arg0]      In   光源セット名
'
'戻り値:
'
'注意事項:
'
Public Sub FW_OptSet(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    OptCheckCounter = 0
    
    If Parameter.ArgParameterCount() < 1 Or Parameter.ArgParameterCount > 2 Then
        Err.Raise 9999, "FW_OptSet", "The number of FW_OptSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------
    
    'Call "OptIni" in case the following object "OptCond" is nothing. (OptcondオブジェクトがNothingだったらOptIniをかける)
    Call sub_CheckOptCond
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_ILLUMINATOR
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================
    
    '光源設定
    Call OptSet(Parameter.Arg(1), Parameter.Arg(0))
    
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

Public Sub FW_OptSetAxis(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    OptCheckCounter = 0
    
    If Parameter.ArgParameterCount() < 1 Or Parameter.ArgParameterCount > 2 Then
        Err.Raise 9999, "FW_OptSet", "The number of FW_OptSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------
    
    'Call "OptIni" in case the following object "OptCond" is nothing. (OptcondオブジェクトがNothingだったらOptIniをかける)
    Call sub_CheckOptCond
    
''=========Before TestCondition Check Start ======================
'    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
'        Dim eMode As eTestCnditionCheck
'        eMode = TCC_ILLUMINATOR
'        Call CheckBeforeTestCondition(eMode, Parameter)
'    End If
''=========Before TestCondition Check End ========================
    
    '光源設定
    Call OptSet_Axis(Parameter.Arg(1), Parameter.Arg(0))
    
''=========After TestCondition Check Start ======================
'    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
'        Call CheckAfterTestCondition(eMode, Parameter)
'    End If
''=========After TestCondition Check End ========================
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

Public Sub FW_OptSetDevice(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    OptCheckCounter = 0
    
    If Parameter.ArgParameterCount() < 1 Or Parameter.ArgParameterCount > 2 Then
        Err.Raise 9999, "FW_OptSet", "The number of FW_OptSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------
    
    'Call "OptIni" in case the following object "OptCond" is nothing. (OptcondオブジェクトがNothingだったらOptIniをかける)
    Call sub_CheckOptCond
    
''=========Before TestCondition Check Start ======================
'    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
'        Dim eMode As eTestCnditionCheck
'        eMode = TCC_ILLUMINATOR
'        Call CheckBeforeTestCondition(eMode, Parameter)
'    End If
''=========Before TestCondition Check End ========================
    
    '光源設定
    Call OptSet_Device(Parameter.Arg(1), Parameter.Arg(0))
    
''=========After TestCondition Check Start ======================
'    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
'        Call CheckAfterTestCondition(eMode, Parameter)
'    End If
''=========After TestCondition Check End ========================
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

Public Sub FW_OptSet_Test(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    OptCheckCounter = 0
    
    If Parameter.ArgParameterCount() < 1 Or Parameter.ArgParameterCount > 2 Then
        Err.Raise 9999, "FW_OptSet", "The number of FW_OptSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------
    
    'Call "OptIni" in case the following object "OptCond" is nothing. (OptcondオブジェクトがNothingだったらOptIniをかける)
    Call sub_CheckOptCond
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_ILLUMINATOR
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================
    
    '光源設定
    Call OptSet_Test(Parameter.Arg(1), Parameter.Arg(0))
    
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

Public Sub FW_OptMod(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 5 Then
        Err.Raise 9999, "FW_OptMod", "The number of FW_OptMod's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    OptCheckCounter = 0
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------

    '光源設定
    With Parameter
        Call OptMod(.Arg(0), .Arg(1))
    End With

    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

Public Sub FW_OptJudgement(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
'    If Parameter.ArgParameterCount() <> 5 Then
'        Err.Raise 9999, "FW_OptMod", "The number of FW_OptMod's arguments is invalid." & " @ " & Parameter.ConditionName
'    End If
'
'    OptCheckCounter = 0

    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------

    Call Opt_Judgment_Test(Parameter.Arg(1)) 'For CIS
'    '光源設定
'    With Parameter
'        Call OptMod(.Arg(0), .Arg(1))
'    End With

    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

Public Sub FW_OptStatus(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
''    Exit Sub
    
    If Parameter.ArgParameterCount() <> 1 Then
        Err.Raise 9999, "FW_OptStatus", "The number of FW_OptMod's arguments is invalid." & " @ " & Parameter.ConditionName
    End If

    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------

    '光源設定
    Dim iStatus As Long
    If OptCond.IllumMaker = NIKON Then
        iStatus = NSIS_II.status
        
        While (iStatus <> 0)
            If OptCheckCounter < 999 Then
                TheHdw.TOPT.Recall
                OptCheckCounter = OptCheckCounter + 1
                Call WaitSet(10 * mS)
                Exit Sub
            End If
            iStatus = NSIS_II.status
        Wend
    End If
'    With Parameter
'        Call OptStatus(.Arg(0))
'    End With

    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

Public Sub FW_OptModZ(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 5 Then
        Err.Raise 9999, "FW_OptModZ", "The number of FW_OptMod's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    OptCheckCounter = 0
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------

    '光源設定
    With Parameter
        Call OptModZ_NSIS5(.Arg(0), .Arg(1))
    End With

    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub
'内容:
'   電源設定を行う
'
'パラメータ:
'    [Arg0]      In   条件名 PowerSuppluyVoltageシートでの名称
'    [Arg1]      In   シーケンス名　PowerSequenceシートでの名称
'    [Arg2]      In   実行後のウェイトタイム(省略可能 省略時Waitなし)
'
'戻り値:
'
'注意事項:
'
Public Sub FW_set_voltage(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 2 And Parameter.ArgParameterCount() <> 3 Then
        Err.Raise 9999, "FW_set_voltage", "The number of FW_set_voltage's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    'CSetFunctionInfoからパラメータの取得
    Dim strPowerVoltageName As String
    Dim strPowerSequenceName As String
    strPowerVoltageName = Parameter.Arg(0)
    strPowerSequenceName = Parameter.Arg(1)
    
    '電圧調整用にPALSの変数に格納。
    Now_Mode = strPowerVoltageName
         
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_SETVOLTAGE
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================
    
    'パタン停止
    Call StopPattern 'EeeJob関数
    
    Const Zero_V_Con As String = "ZERO"
    Const Zero_V_Con2 As String = "ZERO_V"
    
    If UCase(strPowerVoltageName) = Zero_V_Con Or UCase(strPowerVoltageName) = Zero_V_Con2 Then
        '電源Conditionの適用(For ZERO_V) APMU Pinについては、5mAクランプ、50mAレンジに設定する。
        Call ApplyPowerConditionForUS(strPowerVoltageName, strPowerSequenceName)
    Else:
        '電源Conditionの適用
        Call ApplyPowerCondition(strPowerVoltageName, strPowerSequenceName)
    End If

   '待ち時間の指定がある場合、待つ
    If Parameter.ArgParameterCount() = 3 Then
        Dim dblWaitTime As Double
        If Not IsNumeric(Parameter.Arg(2)) Then
            Err.Raise 9999, "FW_set_voltage", "FW_set_voltage's Arg2 is invalid type." & " @ " & Parameter.ConditionName
        End If
        dblWaitTime = Parameter.Arg(2)
        Call TheHdw.WAIT(dblWaitTime)
    End If
    
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================

    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数

End Sub
'内容:
'   ピンの接続を行う
'
'パラメータ:
'    [Arg0]      In   ピン名（ピングループ名）
'
'戻り値:
'
'注意事項:
'       ActiveSiteすべてを実行、サイトシェア禁止
'
Public Sub FW_ConnectIOPins(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 1 Then
        Err.Raise 9999, "FW_ConnectIOPins", "The number of FW_ConnectIOPins's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_TESTER_CHANNELS
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================
    
    Call TheHdw.Digital.relays.Pins(Parameter.Arg(0)).ConnectPins
    
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   ピンの開放を行う
'
'パラメータ:
'    [Arg0]      In   ピン名（ピングループ名）
'    [ArgN-1]
'    [ArgN-1]    In   サイト番号(省略された場合は全サイト)
'
'戻り値:
'
'注意事項:
'
Public Sub FW_DisconnectPins(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount < 1 Then
        Err.Raise 9999, "FW_DisconnectPins", "The number of FW_DisconnectPins's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
        
    Dim strPins As String
    Dim lSite As Long
        
    Dim i As Long
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_TESTER_CHANNELS
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================
    
    strPins = Parameter.Arg(0)
    If Parameter.ArgParameterCount = 1 Then
    
        Call DisconnectPins(strPins)
        
    ElseIf Parameter.ArgParameterCount >= 2 Then
    
        If Parameter.ArgParameterCount > 2 Then
            For i = 1 To Parameter.ArgParameterCount - 2
                strPins = strPins & "," & Parameter.Arg(i)
            Next i
        End If
        
        With Parameter
            If IsNumeric(.Arg(.ArgParameterCount - 1)) Then
                lSite = .Arg(.ArgParameterCount - 1)
                 Call DisconnectPins(strPins, lSite)
            Else
                strPins = strPins & "," & Parameter.Arg(.ArgParameterCount - 1)
                Call DisconnectPins(strPins)
            End If
        End With
        
    End If
    
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================
                    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub


'内容:
'   指定時間分Waitをする
'
'パラメータ:
'    [Arg0]      In   Wait時間(s)
'
'戻り値:
'
'注意事項:
'     TheHdw.Waitで問答無用に待つ
'
Public Sub FW_WaitSet(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 1 Then
        Err.Raise 9999, "FW_WaitSet", "The number of FW_WaitSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_WaitSet", "FW_WaitSet's Arg0 is invalid type." & " @ " & Parameter.ConditionName
    End If
    
    Dim dblWaitTime As Double
    dblWaitTime = Parameter.Arg(0)
    
    'Waitする
    Call TheHdw.WAIT(dblWaitTime)
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   指定時間分Waitをする
'
'パラメータ:
'    [Arg0]      In   Wait時間(s)
'
'戻り値:
'
'注意事項:
'     TheExec.RunOptions.AutoAcquire
'     に応じてWaitを分ける、呼び出し側がTOPT実行しているか意識する必要がある
'     TOPT実行中でなければ期待した動作はしない
'
Public Sub FW_WaitSetTopt(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 1 Then
        Err.Raise 9999, "FW_WaitSet", "The number of FW_WaitSetTopt's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_WaitSet", "FW_WaitSetTopt's Arg0 is invalid type." & " @ " & Parameter.ConditionName
    End If
    
    Dim dblWaitTime As Double
    dblWaitTime = Parameter.Arg(0)
    
    'Waitする
    If TheExec.RunOptions.AutoAcquire = True Then
        Call TheHdw.TOPT.WAIT(toptTimer, dblWaitTime * 1000)
    Else
        Call TheHdw.WAIT(dblWaitTime)
    End If
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub


'内容:
'   ピンにFVMIの設定を行う
'
'パラメータ:
'    [Arg0]      In   ピン名（ピングループ名）
'    [Arg1]      In   フォース電圧
'    [Arg2]      In   クランプ電流
'    [Arg3]      In   サイト番号（-1は省略として扱う）
'    [Arg4]      In   コネクトするかどうか(省略可能：省略時コネクト)
'                       （Falseでコネクトしない、それ以外はコネクト）
'
'戻り値:
'
'注意事項:
'
Public Sub FW_SetFVMI(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 3 And Parameter.ArgParameterCount <> 4 _
        And Parameter.ArgParameterCount <> 5 Then
            Err.Raise 9999, "FW_SetFVMI", "The number of FW_SetFVMI's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPins As String
    Dim dblForce As Double
    Dim dblClamp As Double
    Dim lSite As Long
    Dim strIsConnect As String
    
    strPins = Parameter.Arg(0)
    
    If Not IsNumeric(Parameter.Arg(1)) Then
        Err.Raise 9999, "FW_SetFVMI", "FW_SetFVMI 's Arg1 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblForce = Parameter.Arg(1)
    
    If Not IsNumeric(Parameter.Arg(2)) Then
        Err.Raise 9999, "FW_SetFVMI", "FW_SetFVMI 's Arg2 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblClamp = Parameter.Arg(2)
    
    If Parameter.ArgParameterCount > 3 Then
        If Not IsNumeric(Parameter.Arg(3)) Then
            Err.Raise 9999, "FW_SetFVMI", "FW_SetFVMI 's Arg3 is invalid type." & " @ " & Parameter.ConditionName
        End If
        lSite = Parameter.Arg(3)
    End If
    
    If Parameter.ArgParameterCount = 5 Then
        strIsConnect = Parameter.Arg(4)
    End If
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_TESTER_CHANNELS
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================

    With Parameter
        Select Case .ArgParameterCount
            Case 3
                Call SetFVMI(strPins, dblForce, dblClamp)
            Case 4
                If lSite = -1 Then
                    Call SetFVMI(strPins, dblForce, dblClamp)
                Else
                    Call SetFVMI(strPins, dblForce, dblClamp, lSite)
                End If
            Case 5
                If strIsConnect = "False" Then
                    If lSite = -1 Then
                       Call SetFVMI(strPins, dblForce, dblClamp, , False)
                    Else
                       Call SetFVMI(strPins, dblForce, dblClamp, lSite, False)
                    End If
                Else
                    If lSite = -1 Then
                       Call SetFVMI(strPins, dblForce, dblClamp)
                    Else
                       Call SetFVMI(strPins, dblForce, dblClamp, lSite)
                    End If
                End If
        End Select
    End With
    
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   ピンにFIMVの設定を行う
'
'パラメータ:
'    [Arg0]      In   ピン名（ピングループ名）
'    [Arg1]      In   フォース電流
'    [Arg2]      In   クランプ電圧
'    [Arg3]      In   サイト番号（-1は省略として扱う）
'    [Arg4]      In   コネクトするかどうか(省略可能：省略時コネクト)
'                       （Falseでコネクトしない、それ以外はコネクト）
'
'戻り値:
'
'注意事項:
'
Public Sub FW_SetFIMV(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 3 And Parameter.ArgParameterCount <> 4 _
        And Parameter.ArgParameterCount <> 5 Then
            Err.Raise 9999, "FW_SetFIMV", "The number of FW_SetFIMV's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPins As String
    Dim dblForce As Double
    Dim dblClamp As Double
    Dim lSite As Long
    Dim strIsConnect As String
    
    strPins = Parameter.Arg(0)
    
    If Not IsNumeric(Parameter.Arg(1)) Then
        Err.Raise 9999, "FW_SetFIMV", "FW_SetFIMV 's Arg1 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblForce = Parameter.Arg(1)
    
    If Not IsNumeric(Parameter.Arg(2)) Then
        If UCase(Parameter.Arg(2)) = "NONE" Then
            dblClamp = 5
        Else:
            Err.Raise 9999, "FW_SetFIMV", "FW_SetFIMV 's Arg2 is invalid type." & " @ " & Parameter.ConditionName
        End If
    Else
        dblClamp = Parameter.Arg(2)
    End If
    
    If Parameter.ArgParameterCount > 3 Then
        If Not IsNumeric(Parameter.Arg(3)) Then
            Err.Raise 9999, "FW_SetFIMV", "FW_SetFIMV 's Arg3 is invalid type." & " @ " & Parameter.ConditionName
        End If
        lSite = Parameter.Arg(3)
    End If
    
    If Parameter.ArgParameterCount = 5 Then
        strIsConnect = Parameter.Arg(4)
    End If
    
    With Parameter
        Select Case .ArgParameterCount
            Case 3
                Call SetFIMV(strPins, dblForce, dblClamp)
            Case 4
                If lSite = -1 Then
                    Call SetFIMV(strPins, dblForce, dblClamp)
                Else
                    Call SetFIMV(strPins, dblForce, dblClamp, lSite)
                End If
            Case 5
                If .Arg(4) = "False" Then
                    If lSite = -1 Then
                       Call SetFIMV(strPins, dblForce, dblClamp, , False)
                    Else
                       Call SetFIMV(strPins, dblForce, dblClamp, lSite, False)
                    End If
                Else
                    If lSite = -1 Then
                       Call SetFIMV(strPins, dblForce, dblClamp)
                    Else
                       Call SetFIMV(strPins, dblForce, dblClamp, lSite)
                    End If
                End If
        End Select
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub


'内容:
'   ピンにFVMIの設定を行う（APMU） PowerDown専用　レンジ：50mA、クランプ：5mA　固定となる
'
'パラメータ:
'    [Arg0]      In   ピン名（ピングループ名）
'    [Arg1]      In   フォース電圧
'    [Arg2]      In   クランプ電流
'    [Arg3]      In   サイト番号（-1は省略として扱う）
'    [Arg4]      In   コネクトするかどうか（Falseでコネクトしない、それ以外はコネクト）
'
'戻り値:
'
'注意事項:
'
Public Sub FW_SetFVMI_APMUoff(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 3 And Parameter.ArgParameterCount <> 4 _
        And Parameter.ArgParameterCount <> 5 Then
            Err.Raise 9999, "FW_SetFVMI_APMUoff", "The number of FW_SetFVMI_APMUoff's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPins As String
    Dim dblForce As Double
    Dim dblClamp As Double
    Dim lSite As Long
    Dim strIsConnect As String
    
    strPins = Parameter.Arg(0)
    
    If Not IsNumeric(Parameter.Arg(1)) Then
        Err.Raise 9999, "FW_SetFVMI_APMUoff", "FW_SetFVMI_APMUoff 's Arg1 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblForce = Parameter.Arg(1)
    
    If Not IsNumeric(Parameter.Arg(2)) Then
        Err.Raise 9999, "FW_SetFVMI_APMUoff", "FW_SetFVMI_APMUoff 's Arg2 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblClamp = Parameter.Arg(2)
    
    If Parameter.ArgParameterCount > 3 Then
        If Not IsNumeric(Parameter.Arg(3)) Then
            Err.Raise 9999, "FW_SetFVMI_APMUoff", "FW_SetFVMI_APMUoff 's Arg3 is invalid type." & " @ " & Parameter.ConditionName
        End If
        lSite = Parameter.Arg(3)
    End If
    
    If Parameter.ArgParameterCount = 5 Then
        strIsConnect = Parameter.Arg(4)
    End If
    
    With Parameter
        Select Case .ArgParameterCount
            Case 3
                Call SetFVMI_APMUoff(strPins, dblForce, dblClamp)
            Case 4
                If lSite = -1 Then
                    Call SetFVMI_APMUoff(strPins, dblForce, dblClamp)
                Else
                    Call SetFVMI_APMUoff(strPins, dblForce, dblClamp, lSite)
                End If
            Case 5
                If strIsConnect = "False" Then
                    If lSite = -1 Then
                       Call SetFVMI_APMUoff(strPins, dblForce, dblClamp, , False)
                    Else
                       Call SetFVMI_APMUoff(strPins, dblForce, dblClamp, lSite, False)
                    End If
                Else
                    If lSite = -1 Then
                       Call SetFVMI_APMUoff(strPins, dblForce, dblClamp)
                    Else
                       Call SetFVMI_APMUoff(strPins, dblForce, dblClamp, lSite)
                    End If
                End If
        End Select
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub


'内容:
'   ピンにFVMIの設定を行う（BPMU）
'
'パラメータ:
'    [Arg0]      In   ピン名（ピングループ名）
'    [Arg1]      In   フォース電圧
'    [Arg2]      In   クランプ電流
'    [Arg3]      In   サイト番号（-1は省略として扱う）
'    [Arg4]      In   コネクトするかどうか（Falseでコネクトしない、それ以外はコネクト）
'
'戻り値:
'
'注意事項:
'
Public Sub FW_SetFVMI_BPMU(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 3 And Parameter.ArgParameterCount <> 4 _
        And Parameter.ArgParameterCount <> 5 Then
            Err.Raise 9999, "FW_SetFVMI_BPMU", "The number of FW_SetFVMI_BPMU's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPins As String
    Dim dblForce As Double
    Dim dblClamp As Double
    Dim lSite As Long
    Dim strIsConnect As String
    
    strPins = Parameter.Arg(0)
    
    If Not IsNumeric(Parameter.Arg(1)) Then
        Err.Raise 9999, "FW_SetFVMI_BPMU", "FW_SetFVMI_BPMU 's Arg1 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblForce = Parameter.Arg(1)
    
    If Not IsNumeric(Parameter.Arg(2)) Then
        Err.Raise 9999, "FW_SetFVMI_BPMU", "FW_SetFVMI_BPMU 's Arg2 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblClamp = Parameter.Arg(2)
    
    If Parameter.ArgParameterCount > 3 Then
        If Not IsNumeric(Parameter.Arg(3)) Then
            Err.Raise 9999, "FW_SetFVMI_BPMU", "FW_SetFVMI_BPMU 's Arg3 is invalid type." & " @ " & Parameter.ConditionName
        End If
        lSite = Parameter.Arg(3)
    End If
    
    If Parameter.ArgParameterCount = 5 Then
        strIsConnect = Parameter.Arg(4)
    End If
    
    With Parameter
        Select Case .ArgParameterCount
            Case 3
                Call SetFVMI_BPMU(strPins, dblForce, dblClamp)
            Case 4
                If lSite = -1 Then
                    Call SetFVMI_BPMU(strPins, dblForce, dblClamp)
                Else
                    Call SetFVMI_BPMU(strPins, dblForce, dblClamp, lSite)
                End If
            Case 5
                If strIsConnect = "False" Then
                    If lSite = -1 Then
                       Call SetFVMI_BPMU(strPins, dblForce, dblClamp, , False)
                    Else
                       Call SetFVMI_BPMU(strPins, dblForce, dblClamp, lSite, False)
                    End If
                Else
                    If lSite = -1 Then
                       Call SetFVMI_BPMU(strPins, dblForce, dblClamp)
                    Else
                       Call SetFVMI_BPMU(strPins, dblForce, dblClamp, lSite)
                    End If
                End If
        End Select
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   ピンにFIMVの設定を行う(BPMU)
'
'パラメータ:
'    [Arg0]      In   ピン名（ピングループ名）
'    [Arg1]      In   フォース電流
'    [Arg2]      In   クランプ電圧
'    [Arg3]      In   サイト番号（-1は省略として扱う）
'    [Arg4]      In   コネクトするかどうか(省略可能：省略時コネクト)
'                       （Falseでコネクトしない、それ以外はコネクト）
'
'戻り値:
'
'注意事項:
'
Public Sub FW_SetFIMV_BPMU(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 3 And Parameter.ArgParameterCount <> 4 _
        And Parameter.ArgParameterCount <> 5 Then
            Err.Raise 9999, "FW_SetFIMV_BPMU", "The number of FW_SetFIMV_BPMU's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPins As String
    Dim dblForce As Double
    Dim dblClamp As Double
    Dim lSite As Long
    Dim strIsConnect As String
    
    strPins = Parameter.Arg(0)
    
    If Not IsNumeric(Parameter.Arg(1)) Then
        Err.Raise 9999, "FW_SetFIMV_BPMU", "FW_SetFIMV_BPMU 's Arg1 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblForce = Parameter.Arg(1)
    
    If Not IsNumeric(Parameter.Arg(2)) Then
        Err.Raise 9999, "FW_SetFIMV_BPMU", "FW_SetFIMV_BPMU 's Arg2 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblClamp = Parameter.Arg(2)
    
    If Parameter.ArgParameterCount > 3 Then
        If Not IsNumeric(Parameter.Arg(3)) Then
            Err.Raise 9999, "FW_SetFIMV_BPMU", "FW_SetFIMV_BPMU 's Arg3 is invalid type." & " @ " & Parameter.ConditionName
        End If
        lSite = Parameter.Arg(3)
    End If
    
    If Parameter.ArgParameterCount = 5 Then
        strIsConnect = Parameter.Arg(4)
    End If
    
    With Parameter
        Select Case .ArgParameterCount
            Case 3
                Call SetFIMV_BPMU(strPins, dblForce, dblClamp)
            Case 4
                If lSite = -1 Then
                    Call SetFIMV_BPMU(strPins, dblForce, dblClamp)
                Else
                    Call SetFIMV_BPMU(strPins, dblForce, dblClamp, lSite)
                End If
            Case 5
                If .Arg(4) = "False" Then
                    If lSite = -1 Then
                       Call SetFIMV_BPMU(strPins, dblForce, dblClamp, , False)
                    Else
                       Call SetFIMV_BPMU(strPins, dblForce, dblClamp, lSite, False)
                    End If
                Else
                    If lSite = -1 Then
                       Call SetFIMV_BPMU(strPins, dblForce, dblClamp)
                    Else
                       Call SetFIMV_BPMU(strPins, dblForce, dblClamp, lSite)
                    End If
                End If
        End Select
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   ピンの開放を行う(BPMU)
'
'パラメータ:
'    [Arg0]      In   ピン名（ピングループ名）
'    [Arg3]      In   サイト番号（-1は省略として扱う）
'
'戻り値:
'
'注意事項:
'
Public Sub FW_DisconnectPins_BPMU(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 1 And Parameter.ArgParameterCount <> 2 Then
        Err.Raise 9999, "DisconnectPins_BPMU", "The number of DisconnectPins_BPMU's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
        
    Dim strPins As String
    Dim lSite As Long
        
    strPins = Parameter.Arg(0)
    
    If Parameter.ArgParameterCount = 2 Then
        If Not IsNumeric(Parameter.Arg(1)) Then
            Err.Raise 9999, "DisconnectPins_BPMU", "DisconnectPins_BPMU 's Arg1 is invalid type." & " @ " & Parameter.ConditionName
        End If
        lSite = Parameter.Arg(1)
    End If
    
    If Parameter.ArgParameterCount = 1 Then
        Call DisconnectPins(strPins)
    ElseIf Parameter.ArgParameterCount = 2 Then
        Call DisconnectPins(strPins, lSite)
    End If
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   パタンの開始をおこなう(終了をまたない)
'
'パラメータ:
'    [Arg0]      In   パタン名
'    [Arg1]      In   TSBシート名
'    [Arg2]      In   実行後のウェイトタイム(省略可能 省略時Waitなし)
'
'戻り値:
'
'注意事項:IP750 or Decoder Patは、専用で設定する。
'
Public Sub FW_PatSet(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
'    Const PAT_START_LABEL As String = "pat_start"
    Const PAT_START_LABEL As String = "START"

    If Parameter.ArgParameterCount <> 2 And Parameter.ArgParameterCount <> 3 Then
         Err.Raise 9999, "FW_PatSet", "The number of FW_PatSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPatGroupName As String
    Dim strTsbName As String
        
    With Parameter
        strPatGroupName = .Arg(0)
        strTsbName = .Arg(1)
    End With
            
    Call StopPattern_Halt 'EeeJob関数
    Call SetTimeOut 'EeeJob関数
    
    '分割レジスタ対応ルーチン Start
    'レジスタ設定部Only(keep_alive)：PatRun
    'レジスタ設定+駆動部:PatSet
    Dim tmpPatGroupName() As String
    Dim i As Integer
    tmpPatGroupName = Split(strPatGroupName, ",")
    
    PatCheckCounter = 0
    
    For i = 0 To UBound(tmpPatGroupName)
        If i < UBound(tmpPatGroupName) Then
            With TheHdw.Digital
                Call .Timing.Load(strTsbName)
                Call .Patterns.pat(tmpPatGroupName(i)).Run(PAT_START_LABEL)
            End With
            If TheExec.RunOptions.AutoAcquire = True Then
                Dim iStatus As Long
                If TheHdw.Digital.Patgen.IsRunningAnySite = True Then      'True:Still Running
                    iStatus = 0
                ElseIf TheHdw.Digital.Patgen.IsRunningAnySite = False Then 'False:haltend or keepalive
                    iStatus = 1
                End If
                
                While (iStatus <> 1)
                    If PatCheckCounter < 999 Then
                        TheHdw.TOPT.Recall
                        PatCheckCounter = PatCheckCounter + 1
                        Call WaitSet(10 * mS)
                        Exit Sub
                    End If
                Wend
            End If
        ElseIf i = UBound(tmpPatGroupName) Then
            With TheHdw.Digital
                Call .Timing.Load(strTsbName)
                Call .Patterns.pat(tmpPatGroupName(i)).Start(PAT_START_LABEL)
                    If Flg_Scrn = 1 And tmpPatGroupName(i) = "PG_CUR_SCR" Then
                        Dim Hsn(nSite) As Double
                        Dim site As Long
                            TheHdw.WAIT 50 * mS
                            Call MeasureI_APMU("P_HSN", Hsn, 50)
                        TheResult.Add "IDDBI_HSN", Hsn
                    End If
            End With
        End If
    Next i
    '分割レジスタ対応ルーチン End
    
   '待ち時間の指定がある場合、待つ
    If Parameter.ArgParameterCount() = 3 Then
        Dim dblWaitTime As Double
        If Not IsNumeric(Parameter.Arg(2)) Then
            Err.Raise 9999, "FW_PatSet", "FW_PatSet's Arg2 is invalid type." & " @ " & Parameter.ConditionName
        End If
        dblWaitTime = Parameter.Arg(2)
        Call TheHdw.WAIT(dblWaitTime)
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   パタンの開始をおこなう(終了をまたない)
'
'パラメータ:
'    [Arg0]      In   パタン名
'    [Arg1]      In   TSBシート名
'    [Arg2]      In   実行後のウェイトタイム(省略可能 省略時Waitなし)
'
'戻り値:
'
'注意事項:
'
Public Sub FW_PatSet_Decoder(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
    Const PAT_START_LABEL As String = "pat_start"

    If Parameter.ArgParameterCount <> 2 And Parameter.ArgParameterCount <> 3 Then
         Err.Raise 9999, "FW_PatSet", "The number of FW_PatSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPatGroupName As String
    Dim strTsbName As String
        
    With Parameter
        strPatGroupName = .Arg(0)
        strTsbName = .Arg(1)
    End With
    
    Call StopPattern 'EeeJob関数
    Call SetTimeOut 'EeeJob関数
    
    With TheHdw.Digital
        Call .Timing.Load(strTsbName)
        Call .Patterns.pat(strPatGroupName).Start(PAT_START_LABEL)
    End With
    
   '待ち時間の指定がある場合、待つ
    If Parameter.ArgParameterCount() = 3 Then
        Dim dblWaitTime As Double
        If Not IsNumeric(Parameter.Arg(2)) Then
            Err.Raise 9999, "FW_PatSet", "FW_PatSet's Arg2 is invalid type." & " @ " & Parameter.ConditionName
        End If
        dblWaitTime = Parameter.Arg(2)
        Call TheHdw.WAIT(dblWaitTime)
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   パタンの開始をおこなう(終了をまつ)
'
'パラメータ:
'    [Arg0]      In   パタン名
'    [Arg1]      In   TSBシート名
'    [Arg2]      In   実行後のウェイトタイム(省略可能 省略時Waitなし)
'
'戻り値:
'
'注意事項:
'
Public Sub FW_PatRun(ByVal Parameter As CSetFunctionInfo)

    
    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
'    Const PAT_START_LABEL As String = "pat_start"
    Const PAT_START_LABEL As String = "START"

    If Parameter.ArgParameterCount <> 2 And Parameter.ArgParameterCount <> 3 Then
         Err.Raise 9999, "FW_PatSet", "The number of FW_PatSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPatGroupName As String
    Dim strTsbName As String
        
    With Parameter
        strPatGroupName = .Arg(0)
        strTsbName = .Arg(1)
    End With
    
    Call StopPattern_Halt 'EeeJob関数
    Call SetTimeOut 'EeeJob関数
    
    With TheHdw.Digital
        Call .Timing.Load(strTsbName)
        Call .Patterns.pat(strPatGroupName).Run(PAT_START_LABEL)
    End With
    
   '待ち時間の指定がある場合、待つ
    If Parameter.ArgParameterCount() = 3 Then
        Dim dblWaitTime As Double
        If Not IsNumeric(Parameter.Arg(2)) Then
            Err.Raise 9999, "FW_PatSet", "FW_PatSet's Arg2 is invalid type." & " @ " & Parameter.ConditionName
        End If
        dblWaitTime = Parameter.Arg(2)
        Call TheHdw.WAIT(dblWaitTime)
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   パタンの開始をおこなう(終了をまつ)
'
'パラメータ:
'    [Arg0]      In   パタン名
'    [Arg1]      In   TSBシート名
'    [Arg2]      In   実行後のウェイトタイム(省略可能 省略時Waitなし)
'
'戻り値:
'
'注意事項:
'
Public Sub FW_PatRun_Decoder(ByVal Parameter As CSetFunctionInfo)

    
    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
    Const PAT_START_LABEL As String = "pat_start"

    If Parameter.ArgParameterCount <> 2 And Parameter.ArgParameterCount <> 3 Then
         Err.Raise 9999, "FW_PatSet", "The number of FW_PatSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPatGroupName As String
    Dim strTsbName As String
        
    With Parameter
        strPatGroupName = .Arg(0)
        strTsbName = .Arg(1)
    End With
    
    Call StopPattern 'EeeJob関数
    Call SetTimeOut 'EeeJob関数
    
    With TheHdw.Digital
        Call .Timing.Load(strTsbName)
        Call .Patterns.pat(strPatGroupName).Run(PAT_START_LABEL)
    End With
    
   '待ち時間の指定がある場合、待つ
    If Parameter.ArgParameterCount() = 3 Then
        Dim dblWaitTime As Double
        If Not IsNumeric(Parameter.Arg(2)) Then
            Err.Raise 9999, "FW_PatSet", "FW_PatSet's Arg2 is invalid type." & " @ " & Parameter.ConditionName
        End If
        dblWaitTime = Parameter.Arg(2)
        Call TheHdw.WAIT(dblWaitTime)
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   パタンを停止する
'
'パラメータ:
'    なし
'
'戻り値:
'
'注意事項:
'   パラメータは書いても無視される
'
Public Sub FW_StopPat(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
    Call StopPattern 'EeeJob関数
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   ピンの初期設定を行う
'
'パラメータ:
'    [Arg0]      In   ピン名
'    [Arg1]      In   InitState[Hi, Lo, Off]
'    [Arg2]      In   StartState[Hi, Lo, Off]
'
'戻り値:
'
'注意事項:
'
Public Sub FW_SetIOPinState(ByVal Parameter As CSetFunctionInfo)

On Error GoTo ErrHandler

    If Parameter.ArgParameterCount <> 3 Then
         Err.Raise 9999, "FW_SetIOPinState", "The number of FW_SetIOPinState's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPinName As String
    Dim eInitState As ChInitState
    Dim eStartState As chStartState
    
    With Parameter
        strPinName = .Arg(0)
        eInitState = ConvertInitState(.Arg(1))
        eStartState = ConvertStartState(.Arg(2))
    End With
    
    With TheHdw.Pins(strPinName)
        .InitState = eInitState
        .StartState = eStartState
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

Private Function ConvertInitState(ByVal Arg As String) As ChInitState
    
    Select Case Arg
        Case "chInitHi"
            ConvertInitState = chInitHi
        Case "chInitLo"
            ConvertInitState = chInitLo
        Case "chInitOff"
            ConvertInitState = chInitOff
        Case Else
            Err.Raise 9999, "ConvertInitState", "Init State invalide param" '呼び出し元でエラーハンドリングをしてほしい
    End Select
       
End Function

Private Function ConvertStartState(ByVal Arg As String) As chStartState
    
    Select Case Arg
        Case "chStartHi"
            ConvertStartState = chStartHi
        Case "chStartLo"
            ConvertStartState = chStartLo
        Case "chStartOff"
            ConvertStartState = chStartOff
        Case "chStartFmt"
            ConvertStartState = chStartFmt
        Case "chStartNone"
            ConvertStartState = chStartNone
        Case Else
            Err.Raise 9999, "ConvertStartState", "Start State invalide param" '呼び出し元でエラーハンドリングをしてほしい
    End Select
        
End Function



'内容:
'   ピンの初期設定を行う
'
'パラメータ:
'    [Arg0]      In   ClockVoltageの条件名
'    [Arg1]      In   対象ピンの名前（ClockVoltageシートの名称と等しいこと）
'
'戻り値:
'
'注意事項:
'       ActiveSiteすべてを実行、サイトシェア禁止
'
Public Sub FW_SetIOPinElectronics(ByVal Parameter As CSetFunctionInfo)

On Error GoTo ErrHandler

    If Parameter.ArgParameterCount <> 2 Then
         Err.Raise 9999, "FW_SetIOPinElectronics", "The number of FW_SetIOPinElectronics's arguments is invalid." & " @ " & Parameter.ConditionName
    End If

'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_TESTER_CHANNELS
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================
    
    With Parameter
        Call ShtClockV.GetClockInfo(.Arg(0), .Arg(1)).ForceGroupPins(.Arg(1))
    End With
    
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================

    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub


'内容:
'   ユーザマクロを引数無しでコールする（もしあれば）
'
'パラメータ:
'    [Arg0]      In   ユーザーマクロ名
'
'戻り値:
'
'注意事項:

Public Sub FW_CallUserMacro(ByVal Parameter As CSetFunctionInfo)
    
On Error GoTo ErrHandler

    If Parameter.ArgParameterCount <> 1 Then
         Err.Raise 9999, "FW_CallUserMacro", "The number of FW_CallUserMacro's arguments is invalid." & " @ " & Parameter.ConditionName
    End If

    Call Application.Run(Parameter.Arg(0))
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Serial_Test用:Start
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'内容:
'   SUB電流測定のパラメータ設定を行う
'
'パラメータ:
'    [Arg0]      In   平均回数
'    [Arg1]      In   クランプ電流(A)
'    [Arg2]      In   WaitTime(s)
'    [Arg3]      In   ピンリソース
'
'戻り値:
'
'注意事項:
'2012/10/19 DC_WG 修正
'2013/02/01 MB_WG 修正
'2013/09/27 変更
'
Public Sub FW_SetSubCurrentParam(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 4 Then
         Err.Raise 9999, "FW_SetSubCurrentParam", "The number of FW_SetSubCurrentParam's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '========Check Average Count============================================
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_SetSubCurrentParam", "FW_SetSubCurrentParam Arg0: Type Mismatch """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    Dim lCount As Long
    lCount = Parameter.Arg(0)
    
    '========Check Clamp Current============================================
    If Not IsNumeric(Parameter.Arg(1)) Then
        Err.Raise 9999, "FW_SetSubCurrentParam", "FW_SetSubCurrentParam Arg1: Type Mismatch """ & Parameter.Arg(1) & """ @ " & Parameter.ConditionName
    End If
    
    Dim dblClampCurrent As Double
    dblClampCurrent = Parameter.Arg(1)
    
    '========Check Wait Time ===============================================
    If Not IsNumeric(Parameter.Arg(2)) Then
        Err.Raise 9999, "FW_SetSubCurrentParam", "FW_SetSubCurrentParam Arg2: Type Mismatch """ & Parameter.Arg(2) & """ @ " & Parameter.ConditionName
    End If
    
    Dim dblWaitTime As Double
    dblWaitTime = Parameter.Arg(2)
    
    Dim strResourceName As String
    If UCase(Parameter.Arg(3)) = "BPMU" Then
        strResourceName = "BPMU"
    Else
        strResourceName = "Not BPMU"
    End If
    
    '========Add SubCurrentParam To VarBank====================================
    Dim strCountKey As String, strClampKey As String, strWaitTimeKey As String, strPinResourceKey As String
    strCountKey = GetInstanceName & SUBCURRENT_AVERAGE_COUNT
    strClampKey = GetInstanceName & SUBCURRENT_CLAMP_CURRENT
    strWaitTimeKey = GetInstanceName & SUBCURRENT_WAIT_TIME
    strPinResourceKey = GetInstanceName & SUBCURRENT_PIN_RESOURCE

    With TheVarBank
        If .IsExist(strCountKey) = True Then
            Err.Raise 9999, "FW_SetSubCurrentParam", "FW_SetSubCurrentParam was already called for AverageCount: "
        ElseIf .IsExist(strClampKey) = True Then
            Err.Raise 9999, "FW_SetSubCurrentParam", "FW_SetSubCurrentParam was already called for ClampCurrent: "
        ElseIf .IsExist(strWaitTimeKey) = True Then
            Err.Raise 9999, "FW_SetSubCurrentParam", "FW_SetSubCurrentParam was already called for WaitTime: "
        ElseIf .IsExist(strPinResourceKey) = True Then
            Err.Raise 9999, "FW_SetSubCurrentParam", "FW_SetSubCurrentParam was already called for PinResource: "
        Else
            Call .Add(strCountKey, lCount, False, strCountKey)
            Call .Add(strClampKey, dblClampCurrent, False, strClampKey)
            Call .Add(strWaitTimeKey, dblWaitTime, False, strWaitTimeKey)
            Call .Add(strPinResourceKey, strResourceName, False, strPinResourceKey)
        End If
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub
'内容:
'   SUB電流測定のパラメータ取得
'
'パラメータ:
'
'戻り値:
'       平均回数
'
'注意事項:
'
Public Function GetSubCurrentAverageCount(ByVal strInstanceName As String) As Long

    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & SUBCURRENT_AVERAGE_COUNT) Then
        Err.Raise 9999, "GetSubCurrentAverageCount", "FW_SetSubCurrentParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetSubCurrentAverageCount = TheVarBank.Value(strInstanceName & SUBCURRENT_AVERAGE_COUNT)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Function

'内容:
'   SUB電流測定のパラメータ取得
'
'パラメータ:
'
'戻り値:
'       クランプ電流値(A)
'
'注意事項:
'
Public Function GetSubCurrentClampCurrent(ByVal strInstanceName As String) As Double

    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & SUBCURRENT_CLAMP_CURRENT) Then
        Err.Raise 9999, "GetSubCurrentClampCurrent", "FW_SetSubCurrentParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetSubCurrentClampCurrent = TheVarBank.Value(strInstanceName & SUBCURRENT_CLAMP_CURRENT)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数

End Function

'内容:
'   SUB電流測定のパラメータ取得
'
'パラメータ:
'
'戻り値:
'       WaitTime(s)
'
'注意事項:
'
Public Function GetSubCurrentWaitTime(ByVal strInstanceName As String) As Double

    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & SUBCURRENT_WAIT_TIME) Then
        Err.Raise 9999, "GetSubCurrentWaitTime", "FW_SetSubCurrentParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetSubCurrentWaitTime = TheVarBank.Value(strInstanceName & SUBCURRENT_WAIT_TIME)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数

End Function
'内容:
'   キャプチャパラメータの設定を行う
'
'パラメータ:
'    [Arg0]      In   平均回数
'    [Arg1]      In   平均化モード(Average or NoAverage)
'
'戻り値:
'
'注意事項:
'
'2013/01/22 H.Arikawa Arg3 -> Arg21へ変更
Public Sub FW_SetCaptureParam(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler:
    
    '========Param check  ==================================================
    If Parameter.ArgParameterCount() <> 4 Then
        Err.Raise 9999, "FW_SetCaptureParam", "The Number of arguments is invalid! """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    
    '========Check Average Count============================================
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_SetCaptureParam", "Arg0: Type Mismatch """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    Dim lCount As Long
    lCount = Parameter.Arg(0)
    If (lCount < 1) Or (512 < lCount) Then 'Check For CaptureUnit
        Err.Raise 9999, "FW_SetCaptureParam", "Arg0: Range Invalid """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    '========Check Average Mode============================================
    Dim strAverageMode As String
    strAverageMode = Parameter.Arg(1)
    If lCount = 1 And strAverageMode = "NonAverage" Then strAverageMode = "Average"
    
    If strAverageMode <> CAPTURE_AVERAGE_MODE_AVERAGE And strAverageMode <> CAPTURE_AVERAGE_MODE_NO_AVERAGE Then
        Err.Raise 9999, "FW_SetCaptureParam", " Arg1: Value Invalid  """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    '========Check Frame Skip Count============================================
    If Not IsNumeric(Parameter.Arg(2)) Then
        Err.Raise 9999, "FW_SetCaptureParam", "Arg2: Type Mismatch """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    Dim lSkipCount As Long
    lSkipCount = Parameter.Arg(2)
    If (lCount < 0) Or (512 < lCount) Then 'Check For CaptureUnit
        Err.Raise 9999, "FW_SetCaptureParam", "Arg2: Range Invalid [0-512] """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    Dim acqInstances() As String
    acqInstances = Split(Parameter.Arg(3), ",")
    Dim i As Long
    '========Add Average Set To VarBank====================================
    Dim strCountKey As String, strModekey As String, strSkipKey As String
    
    For i = 0 To UBound(acqInstances)
        strCountKey = acqInstances(i) & CAPTURE_PARAM_AVERAGE_COUNT
        strModekey = acqInstances(i) & CAPTURE_PARAM_AVERAGE_MODE
        strSkipKey = acqInstances(i) & CAPTURE_PARAM_FRAME_SKIP
    
        With TheVarBank
            If .IsExist(strCountKey) = True Then
                If .Value(strCountKey) <> lCount Then
                    Err.Raise 9999, "FW_SetCaptureParam", "FW_SetCaptureParam already called for AverageCount  """ & .Value(strCountKey) & """ @ " & Parameter.ConditionName
                Else
                    'なにもしない
                End If
            Else
                Call .Add(strCountKey, lCount, False, strCountKey)
            End If
            
            If .IsExist(strModekey) = True Then
                If .Value(strModekey) <> strAverageMode Then
                    Err.Raise 9999, "FW_SetCaptureParam", "FW_SetCaptureParam already called for AverageMode  """ & .Value(strCountKey) & """ @ " & Parameter.ConditionName
                    'なにもしない
                End If
            Else
                Call .Add(strModekey, strAverageMode, False, strModekey)
            End If
            
            If .IsExist(strSkipKey) = True Then
                If .Value(strSkipKey) <> lSkipCount Then
                    Err.Raise 9999, "FW_SetCaptureParam", "FW_SetCaptureParam already called for FrameSkip  """ & .Value(strCountKey) & """ @ " & Parameter.ConditionName
                    'なにもしない
                End If
            Else
                Call .Add(strSkipKey, lSkipCount, False, strSkipKey)
            End If
            
        End With
    Next i
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   画像キャプチャのパラメータ取得
'
'パラメータ:
'
'戻り値:
'       平均回数
'
'注意事項:
'
Public Function GetCaptureParamAverageCount(ByVal strInstanceName As String) As Long
    
    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & CAPTURE_PARAM_AVERAGE_COUNT) Then
        Err.Raise 9999, "GetCaptureParamAverageCount", "FW_SetCaptureParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetCaptureParamAverageCount = TheVarBank.Value(strInstanceName & CAPTURE_PARAM_AVERAGE_COUNT)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Function

'内容:
'   画像キャプチャのパラメータ取得
'
'パラメータ:
'
'戻り値:
'       平均化モード
'
'注意事項:
'
Public Function GetCaptureParamAverageMode(ByVal strInstanceName As String) As IdpAverageMode
    
    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & CAPTURE_PARAM_AVERAGE_MODE) Then
        Err.Raise 9999, "GetCaptureParamAverageMode", "FW_SetCaptureParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    With TheVarBank
        If (.Value(strInstanceName & CAPTURE_PARAM_AVERAGE_MODE) = CAPTURE_AVERAGE_MODE_AVERAGE) Then
            GetCaptureParamAverageMode = idpAverage
            Exit Function
        End If
         If (.Value(strInstanceName & CAPTURE_PARAM_AVERAGE_MODE) = CAPTURE_AVERAGE_MODE_NO_AVERAGE) Then
            GetCaptureParamAverageMode = idpNonAverage
            Exit Function
        End If
   End With
   
   Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Function

'内容:
'   画像キャプチャのパラメータ取得
'
'パラメータ:
'
'戻り値:
'       FrameSkip数
'
'注意事項:
'
Public Function GetCaptureParamFrameSkip(ByVal strInstanceName As String) As Long
    
    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & CAPTURE_PARAM_FRAME_SKIP) Then
        Err.Raise 9999, "GetCaptureParamFrameSkip", "FW_SetCaptureParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetCaptureParamFrameSkip = TheVarBank.Value(strInstanceName & CAPTURE_PARAM_FRAME_SKIP)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Function

'内容:
'   SCRN時のSetMVからMeasureVまでのWait時間設定
'
'パラメータ:
'    [Arg0]      In   Average Count
'    [Arg1]      In   WaitTime(s)
'
'戻り値:
'
'注意事項:
'
Public Sub FW_SetScrnMeasureParam(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 2 Then
         Err.Raise 9999, "FW_SetScrnMeasureParam", "The number of FW_SetScrnMeasureParam's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '========Average ============================================
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_SetScrnMeasureParam", "FW_SetScrnMeasureParam Arg0: Type Mismatch """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    Dim lAverage As Long
    lAverage = Parameter.Arg(0)
    
    
    '========Wait Time ============================================
    If Not IsNumeric(Parameter.Arg(1)) Then
        Err.Raise 9999, "FW_SetScrnMeasureParam", "FW_SetScrnMeasureParam Arg0: Type Mismatch """ & Parameter.Arg(1) & """ @ " & Parameter.ConditionName
    End If
    
    Dim dblWaitTime As Double
    dblWaitTime = Parameter.Arg(1)
    
    
    '========Add SubCurrentParam To VarBank====================================
    Dim strCountKey As String
    strCountKey = GetInstanceName & SCRN_MEAS_AVERAGE_COUNT
    Dim strWaitTimeKey As String
    strWaitTimeKey = GetInstanceName & SCRN_MEAS_WAIT_TIME

    With TheVarBank
        If .IsExist(strCountKey) = True Then
            Err.Raise 9999, "FW_ScrnMeasureWaitParam", "FW_ScrnMeasureWaitParam was already called for WaitTime: "
        Else
            Call .Add(strCountKey, lAverage, False, strCountKey)
        End If
    End With
    
    With TheVarBank
        If .IsExist(strWaitTimeKey) = True Then
            Err.Raise 9999, "FW_ScrnMeasureWaitParam", "FW_ScrnMeasureWaitParam was already called for WaitTime: "
        Else
            Call .Add(strWaitTimeKey, dblWaitTime, False, strWaitTimeKey)
        End If
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   SCRNのパラメータ取得
'
'パラメータ:
'
'戻り値:
'       SetMVからMeasureまでのWait時間取得
'
'注意事項:
'
Public Function GetScrnMeasureWaitTime(ByVal strInstanceName As String) As Double
    
    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & SCRN_MEAS_WAIT_TIME) Then
        Err.Raise 9999, "GetScrnMeasureWaitParam", "GetScrnMeasureWaitParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetScrnMeasureWaitTime = TheVarBank.Value(strInstanceName & SCRN_MEAS_WAIT_TIME)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Function


'内容:
'   SCRNのパラメータ取得
'
'パラメータ:
'
'戻り値:
'       SCRN時測定の平均回数
'
'注意事項:
'
Public Function GetScrnMeasureAverageCount(ByVal strInstanceName As String) As Double
    
    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & SCRN_MEAS_AVERAGE_COUNT) Then
        Err.Raise 9999, "GetScrnMeasureAverageCount", "GetScrnMeasureAverageCount in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetScrnMeasureAverageCount = TheVarBank.Value(strInstanceName & SCRN_MEAS_AVERAGE_COUNT)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Function
'内容:
'   To execute wait for screening.
'
'パラメータ:
'   [Arg0]  In  Screening wait time in second, specified on the specification sheet.
'   [Arg1]  In  Wait time between "SET" and "MEASUREMENT" for the dc test
'   [Arg1]  In  WaitTime(s)
'
'戻り値:
'
'注意事項:
'
Public Sub FW_ScreeningWait(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    Const NUMBER_OF_ARGUMENTS As Long = 3
    If Parameter.ArgParameterCount <> NUMBER_OF_ARGUMENTS Then
         Err.Raise 9999, "FW_ScreeningWait", "The number of FW_ScreeningWait's arguments must be " & NUMBER_OF_ARGUMENTS & "." & " @ " & Parameter.ConditionName
         GoTo ErrHandler
    End If
    
    '======== specification wait time ============================================
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_ScreeningWait", "FW_ScreeningWait Arg0: Type Mismatch (must be numeric) """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
        GoTo ErrHandler
    End If
    
    Dim dblScreeningWait As Double
    dblScreeningWait = Parameter.Arg(0)
    
    
    '========DC measurement wait Time ============================================
    If Not IsNumeric(Parameter.Arg(1)) Then
        If Parameter.Arg(1) = "-" Then
            Parameter.Arg(1) = 0
        Else
            Err.Raise 9999, "FW_ScreeningWait", "FW_ScreeningWait Arg0: Type Mismatch (must be numeric or " - " ) """ & Parameter.Arg(1) & """ @ " & Parameter.ConditionName
            GoTo ErrHandler
        End If
    End If
    
    Dim dblDcWaitTime As Double
    dblDcWaitTime = Parameter.Arg(1)
    
    '========TOPT mode ============================================
    Dim isToptMode As Boolean
    isToptMode = Parameter.Arg(2)
    
    'Waitする
    Dim dblTotalWaitTime As Double
    If dblScreeningWait > dblDcWaitTime Then dblTotalWaitTime = dblScreeningWait - dblDcWaitTime
    If TheExec.RunOptions.AutoAcquire = True And isToptMode Then
        Call TheHdw.TOPT.WAIT(toptTimer, dblTotalWaitTime * 1000)
    Else
        Call TheHdw.WAIT(dblTotalWaitTime)
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub


'内容:
'   ピンをGNDに接続する(接地)
'
'パラメータ:
'    [Arg0]      In   ピン名（ピングループ名）
'
'戻り値:
'
'注意事項:
'       全Siteを1ピンずつGNDに接続。
'
Public Sub FW_SetGND(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 1 Then
        Err.Raise 9999, "FW_SetGND", "The number of FW_SetGND's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_TESTER_CHANNELS
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================
    Call SetGND(Parameter.Arg(0))
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   SUB電流測定のパラメータ取得
'
'パラメータ:
'
'戻り値:
'       PinResourceName
'
'注意事項:
'     2012/10/19 DC_WG 追加
'     2012/11/1  Stop Delete
'
Public Function GetSubCurrentPinResourceName(ByVal strInstanceName As String) As String

    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & SUBCURRENT_PIN_RESOURCE) Then
        Err.Raise 9999, "GetSubCurrentPinResourceName", "FW_SetSubCurrentParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetSubCurrentPinResourceName = TheVarBank.Value(strInstanceName & SUBCURRENT_PIN_RESOURCE)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数

End Function
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Serial_Test用:End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'C撮_OV項目用リレー設定:Start
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'内容:
'   リレーONを行う
'
'パラメータ:
'    [Arg0]      In   リレーUB
'
'戻り値:
'
'注意事項:S撮では使用しないが、弊害はない為、共通のConditionMacroとする為に入れておく。
'     2012/11/1  Stop Delete
'
'''Public Sub FW_SET_RELAY_ON(ByVal Parameter As CSetFunctionInfo)
'''
'''    On Error GoTo ErrHandler
'''
'''    If Parameter.ArgParameterCount() <> 1 Then
'''        err.Raise 9999, "FW_SET_RELAY_ON", "The number of FW_SET_RELAY_ON's arguments is invalid." & " @ " & Parameter.ConditionName
'''    End If
'''
''''=========Before TestCondition Check Start ======================
'''#If EEEAUTO_AUTO_MODIFY_TESTCONDITION = 1 Then
'''    Dim eMode As eTestCnditionCheck
'''    eMode = APMU_RELAY_UB_ON
'''    Call CheckBeforeTestCondition(eMode, Parameter)
'''#End If
''''=========Before TestCondition Check End ========================
'''
'''    'RELAY_ON
'''    DutConnectDbNumber Parameter.Arg(0), True
'''
''''=========After TestCondition Check Start ======================
'''#If EEEAUTO_AUTO_MODIFY_TESTCONDITION = 1 Then
'''    Call CheckAfterTestCondition(eMode, Parameter)
'''#End If
''''=========After TestCondition Check End ========================
'''
'''    Exit Sub
'''
'''ErrHandler:
'''    MsgBox "Error Occured !! " & CStr(err.Number) & " - " & err.Source & chR(13) & chR(13) & err.Description
'''    Call DisableAllTest 'EeeJob関数
'''
'''End Sub

'内容:
'   リレーOFFを行う
'
'パラメータ:
'    [Arg0]      In   リレーUB
'
'戻り値:
'
''''注意事項:S撮では使用しないが、弊害はない為、共通のConditionMacroとする為に入れておく。
''''     2012/11/1  Stop Delete
''''
'''Public Sub FW_SET_RELAY_OFF(ByVal Parameter As CSetFunctionInfo)
'''
'''    On Error GoTo ErrHandler
'''
'''    If Parameter.ArgParameterCount() <> 1 Then
'''        err.Raise 9999, "FW_SET_RELAY_OFF", "The number of FW_SET_RELAY_OFF's arguments is invalid." & " @ " & Parameter.ConditionName
'''    End If
'''
''''=========Before TestCondition Check Start ======================
'''#If EEEAUTO_AUTO_MODIFY_TESTCONDITION = 1 Then
'''    Dim eMode As eTestCnditionCheck
'''    eMode = APMU_RELAY_UB_OFF
'''    Call CheckBeforeTestCondition(eMode, Parameter)
'''#End If
''''=========Before TestCondition Check End ========================
'''
'''    'RELAY_OFF
'''    DutConnectDbNumber Parameter.Arg(0), False
'''
''''=========After TestCondition Check Start ======================
'''#If EEEAUTO_AUTO_MODIFY_TESTCONDITION = 1 Then
'''    Call CheckAfterTestCondition(eMode, Parameter)
'''#End If
''''=========After TestCondition Check End ========================
'''
'''    Exit Sub
'''
'''ErrHandler:
'''    MsgBox "Error Occured !! " & CStr(err.Number) & " - " & err.Source & chR(13) & chR(13) & err.Description
'''    Call DisableAllTest 'EeeJob関数
'''
'''End Sub
'''
'''Public Function DutConnectDbNumber(ByVal dbNum As Long, ByVal ONOFF As Boolean)
'''
'''    TheHdw.APMU.board(APMU_BOARD_NUMBER).UtilityBit(dbNum) = ONOFF '2012/11/15 175Debug Arikawa
'''
'''End Function
'''''''♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'C撮_OV項目用リレー設定:End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'Hold_Voltage_Test用:Start
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'内容:
'   HOLD電圧測定のパラメータ設定を行う
'
'パラメータ:
'    [Arg0]      In   平均回数
'    [Arg1]      In   クランプ電流(A)
'    [Arg2]      In   WaitTime(s)
'
'戻り値:
'
'注意事項:
'     2012/11/1  Stop Delete

Public Sub FW_SetHoldVoltageParam(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 3 Then
         Err.Raise 9999, "FW_SetHoldVoltageParam", "The number of FW_SetHoldVoltageParam's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '========Check Average Count============================================
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_SetHoldVoltageParam", "FW_SetHoldVoltageParam Arg0: Type Mismatch """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    Dim lCount As Long
    lCount = Parameter.Arg(0)
    
    '========Check Clamp Current============================================
    If Not IsNumeric(Parameter.Arg(1)) Then
        Err.Raise 9999, "FW_SetHoldVoltageParam", "FW_SetHoldVoltageParam Arg1: Type Mismatch """ & Parameter.Arg(1) & """ @ " & Parameter.ConditionName
    End If
    
    Dim dblClampCurrent As Double
    dblClampCurrent = Parameter.Arg(1)
    
    '========Check Wait Time ===============================================
    If Not IsNumeric(Parameter.Arg(2)) Then
        Err.Raise 9999, "FW_SetHoldVoltageParam", "FW_SetHoldVoltageParam Arg2: Type Mismatch """ & Parameter.Arg(2) & """ @ " & Parameter.ConditionName
    End If
    
    Dim dblWaitTime As Double
    dblWaitTime = Parameter.Arg(2)
    
    '========Add HoldVoltageParam To VarBank====================================
    Dim strCountKey As String, strClampKey As String, strWaitTimeKey As String
    strCountKey = GetInstanceName & HOLDVOLTAGE_AVERAGE_COUNT
    strClampKey = GetInstanceName & HOLDVOLTAGE_CLAMP_CURRENT
    strWaitTimeKey = GetInstanceName & HOLDVOLTAGE_WAIT_TIME

    With TheVarBank
        If .IsExist(strCountKey) = True Then
            Err.Raise 9999, "FW_SetHoldVoltageParam", "FW_SetHoldVoltageParam was already called for AverageCount: "
        ElseIf .IsExist(strClampKey) = True Then
            Err.Raise 9999, "FW_SetHoldVoltageParam", "FW_SetHoldVoltageParam was already called for ClampCurrent: "
        ElseIf .IsExist(strWaitTimeKey) = True Then
            Err.Raise 9999, "FW_SetHoldVoltageParam", "FW_SetHoldVoltageParam was already called for WaitTime: "
        Else
            Call .Add(strCountKey, lCount, False, strCountKey)
            Call .Add(strClampKey, dblClampCurrent, False, strClampKey)
            Call .Add(strWaitTimeKey, dblWaitTime, False, strWaitTimeKey)
        End If
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   HOLD電圧測定のパラメータ取得
'
'パラメータ:
'
'戻り値:
'       平均回数
'
'注意事項:
'     2012/11/1  Stop Delete

Public Function GetHoldVoltageAverageCount(ByVal strInstanceName As String) As Long

    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & HOLDVOLTAGE_AVERAGE_COUNT) Then
        Err.Raise 9999, "GetHoldVoltageAverageCount", "FW_SetHoldVoltageParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetHoldVoltageAverageCount = TheVarBank.Value(strInstanceName & HOLDVOLTAGE_AVERAGE_COUNT)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Function

'内容:
'   HOLD電圧測定のパラメータ取得
'
'パラメータ:
'
'戻り値:
'       クランプ電流値(A)
'
'注意事項:
'     2012/11/1  Stop Delete

Public Function GetHoldVoltageClampCurrent(ByVal strInstanceName As String) As Double

    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & HOLDVOLTAGE_CLAMP_CURRENT) Then
        Err.Raise 9999, "GetHoldVoltageClampCurrent", "FW_SetHoldVoltageParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetHoldVoltageClampCurrent = TheVarBank.Value(strInstanceName & HOLDVOLTAGE_CLAMP_CURRENT)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数

End Function

'内容:
'   HOLD電圧測定のパラメータ取得
'
'パラメータ:
'
'戻り値:
'       WaitTime(s)
'
'注意事項:
'     2012/11/1  Stop Delete

Public Function GetHoldVoltageWaitTime(ByVal strInstanceName As String) As Double

    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & HOLDVOLTAGE_WAIT_TIME) Then
        Err.Raise 9999, "GetHoldVoltageWaitTime", "FW_SetHoldVoltageParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetHoldVoltageWaitTime = TheVarBank.Value(strInstanceName & HOLDVOLTAGE_WAIT_TIME)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数

End Function
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'Hold_Voltage_Test用:End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'DC TOPT用　FW_DcTopt_Set:
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
'

'内容:
'   DC TOPT利用時に、SET側のシナリオを呼び出し、DC測定の安定化時間分TOPT Waitを呼び出す
'
'パラメータ:
'    [Arg0]      In DC Test Scenario Name。SET側のシナリオ名
'    [Arg1]     DC Test ScenarioのSET側シナリオとMEASURE側シナリオの間の
'               Wait時間。
'
'戻り値:
'
'注意事項:
'     DC TOPTの場合、SETとMEASURE間のWait時間はDC Test Scenarioでは制御せず、
'   ここの値で制御します。従って、デバッグ時には、Test Conditionの本関数
'   のArg1の値を編集してください。
'
Public Sub FW_DcTopt_Set(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 2 Then
        Err.Raise 9999, "FW_DcTopt_Set", "The number of FW_DcTopt_Set's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
                        
    '========== DCシナリオシート実行 ===============================
    TheDcTest.SetScenario (Parameter.Arg(0))
    TheDcTest.Execute
        
    '========= Wait ======================
    If TheExec.RunOptions.AutoAcquire = True Then
        Call TheHdw.TOPT.WAIT(toptTimer, Parameter.Arg(1) * 1000)
    Else
        Call TheHdw.WAIT(Parameter.Arg(1))
    End If
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'DC TOPT用　FW_DcTopt_Measure:
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
'

'内容:
'   DC TOPT利用時に、MEASURE側のシナリオを呼び出す。
'
'パラメータ:
'    [Arg0]      In DC Test Scenario Name。MEASURE側のシナリオ名
'
'戻り値:
'
'注意事項:
'   DC TOPT利用時に、DC Test Scenarioシートで、MEASURE側のシナリオに
'   Waitを記載しても有効にはなりません。これは、DC Test Scenarioシートでは、
'   "MEASURE"アクションのWaitは、同一シナリオの直前に実行されている"SET"
'   アクションからのWait時間として扱われるためです。
'   DC TOPT利用時には、シナリオは"MEASURE"アクションから開始し、直前の"SET"
'   アクションがありませんので、Wait時間を記載しても無視されます。
'
Public Sub FW_DcTopt_Measure(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 1 Then
        Err.Raise 9999, "FW_DcTopt_Measure", "The number of FW_DcTopt_Measure's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
                        
    '========== DCシナリオシート実行 ===============================
    TheDcTest.SetScenario (Parameter.Arg(0))
    TheDcTest.Execute
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'SCRN TOPT用　FW_ScrnWaitSet:
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
'

'内容:
'   指定時間分Waitをする
'
'パラメータ:
'    [Arg0]      In   Wait時間(s)
'
'戻り値:
'
'注意事項:
'     TheHdw.Waitで問答無用に待つ
'     2012/11/1  Stop Delete
'
Public Sub FW_ScrnWaitSet(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 1 Then
        Err.Raise 9999, "FW_ScrnWaitSet", "The number of FW_ScrnWaitSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_ScrnWaitSet", "FW_ScrnWaitSet's Arg0 is invalid type." & " @ " & Parameter.ConditionName
    End If
    
    Dim dblWaitTime As Double
    dblWaitTime = Parameter.Arg(0)
    
    'Waitする
    Call TheHdw.WAIT(dblWaitTime)
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'SCRN TOPT用　FW_ScrnWaitSetTopt:
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
'

'内容:
'   指定時間分Waitをする
'
'パラメータ:
'    [Arg0]      In   Wait時間(s)
'
'戻り値:
'
'注意事項:
'     TheExec.RunOptions.AutoAcquire
'     に応じてWaitを分ける、呼び出し側がTOPT実行しているか意識する必要がある
'     TOPT実行中でなければ期待した動作はしない
'     2012/11/1  Stop Delete
'
Public Sub FW_ScrnWaitSetTopt(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 1 Then
        Err.Raise 9999, "FW_ScrnWaitSetTopt", "The number of FW_ScrnWaitSetTopt's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_ScrnWaitSetTopt", "FW_ScrnWaitSetTopt's Arg0 is invalid type." & " @ " & Parameter.ConditionName
    End If
    
    Dim dblWaitTime As Double
    dblWaitTime = Parameter.Arg(0)
    
    'Waitする
    If TheExec.RunOptions.AutoAcquire = True Then
        Call TheHdw.TOPT.WAIT(toptTimer, dblWaitTime * 1000)
    Else
        Call TheHdw.WAIT(dblWaitTime)
    End If
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub
Public Sub FW_PowerDownAndDisconnectPins(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler

    If Parameter.ArgParameterCount() <> 0 Then
        Err.Raise 9999, "FW_PowerDownAndDisconnectPins", "The number of arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Call PowerDownAndDisconnect
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

Public Sub FW_PowerDown(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler

    If Parameter.ArgParameterCount() <> 0 Then
        Err.Raise 9999, "FW_PowerDown", "The number of arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Call PowerDown4ApmuUnderShoot
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   パターンが走っているかどうかステータスを取得する。
'
'パラメータ:
'
'
'戻り値:
'
'注意事項:Haltを使用しているパターンでのみ使用する。(※無限Loopになる為)
'         keep_alive使用タイプは要確認(12/20)
'GUIでユーザーがTOPT有りのJOB生成を選択した場合にFW_PatSetTypeSelectとセットに生成されるCondition。
'

Public Sub FW_PatStatus(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
    If Parameter.ArgParameterCount() <> 0 Then
        Err.Raise 9999, "FW_PatStatus", "The number of FW_PatStatus's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    If TheExec.RunOptions.AutoAcquire = True Then
        Dim iStatus As Long
        If TheHdw.Digital.Patgen.IsRunningAnySite = True Then      'True:Still Running
            iStatus = 0
        ElseIf TheHdw.Digital.Patgen.IsRunningAnySite = False Then 'False:haltend or keepalive
            iStatus = 1
        End If
        
        While (iStatus <> 1)
            If PatCheckCounter < 999 Then
                TheHdw.TOPT.Recall
                PatCheckCounter = PatCheckCounter + 1
                Call WaitSet(10 * mS)
                Exit Sub
            End If
        Wend
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   パタンの開始をおこなう(終了を待たない/終了を待つ)
'
'パラメータ:
'    [Arg0]      In   パタン名
'    [Arg1]      In   TSBシート名
'    [Arg2]      In   実行後のウェイトタイム(省略可能 省略時Waitなし)
'
'戻り値:
'
'GUIでユーザーがTOPT有りのJOB生成を選択した場合にPatRunの代わりに生成されるCondition。
'
Public Sub FW_PatSetTypeSelect(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
    Const PAT_START_LABEL As String = "START"

    If Parameter.ArgParameterCount <> 2 And Parameter.ArgParameterCount <> 3 Then
         Err.Raise 9999, "FW_PatSet", "The number of FW_PatSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPatGroupName As String
    Dim strTsbName As String
        
    With Parameter
        strPatGroupName = .Arg(0)
        strTsbName = .Arg(1)
    End With
    
    Call StopPattern_Halt 'EeeJob関数
    Call SetTimeOut 'EeeJob関数
    
    'TOPTの使用に応じてPatRun/PatSet動作を選択する。
    If TheExec.RunOptions.AutoAcquire = True Then
        With TheHdw.Digital
            Call .Timing.Load(strTsbName)
            Call .Patterns.pat(strPatGroupName).Start(PAT_START_LABEL)
        End With
    ElseIf TheExec.RunOptions.AutoAcquire = False Then
        With TheHdw.Digital
            Call .Timing.Load(strTsbName)
            Call .Patterns.pat(strPatGroupName).Run(PAT_START_LABEL)
        End With
    End If
    
   '待ち時間の指定がある場合、待つ
    If Parameter.ArgParameterCount() = 3 Then
        Dim dblWaitTime As Double
        If Not IsNumeric(Parameter.Arg(2)) Then
            Err.Raise 9999, "FW_PatSet", "FW_PatSet's Arg2 is invalid type." & " @ " & Parameter.ConditionName
        End If
        dblWaitTime = Parameter.Arg(2)
        Call TheHdw.WAIT(dblWaitTime)
    End If
    
    PatCheckCounter = 0
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'01/29 Add H.Arikawa
'光源の条件設定省略の情報を元に動作を決定する。(光源種類に応じて)
'暫定処理を入れる。

Public Sub FW_OptEscape(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 2 Then
        Err.Raise 9999, "FW_OptEscape", "The number of FW_OptEscape's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------
    
    'Call "OptIni" in case the following object "OptCond" is nothing. (OptcondオブジェクトがNothingだったらOptIniをかける)
    Call sub_CheckOptCond
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_ILLUMINATOR_ESCAPSE
        
        Call CheckBeforeTestCondition(eMode, Parameter)
        If Not IsValidTestCondition(eMode, Parameter) Then Exit Sub
        
    End If
'=========Before TestCondition Check End ========================
    
    OptCheckCounter = 0
        
    '光源設定 Escape
    'NSIS3/3KAI : PINに退避する。
    'NSIS5/5KAI : Upに退避する。
    
    If OptCond.IllumMaker = NIKON Then
        With Parameter
            If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS3KAI" Then
                Call OptMod("PIN", .Arg(0))
            ElseIf OptCond.IllumModel = "N-SIS5" Or OptCond.IllumModel = "N-SIS5KAI" Then
                Call OptModZ_NSIS5("Up", .Arg(0))
            End If
        End With
    End If
    
    Exit Sub
   
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'01/29 Add H.Arikawa
'光源の条件設定省略の情報を元に動作を決定する。(光源種類に応じて)
'暫定処理を入れる。

Public Sub FW_OptModOrModZ1(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 2 Then
        Err.Raise 9999, "FW_OptEscape", "The number of FW_OptEscape's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------
    
    'Call "OptIni" in case the following object "OptCond" is nothing. (OptcondオブジェクトがNothingだったらOptIniをかける)
    Call sub_CheckOptCond
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_ILLUMINATOR_MODZ1
        Call CheckBeforeTestCondition(eMode, Parameter)
        If Not IsValidTestCondition(eMode, Parameter) Then Exit Sub
    
    End If
'=========Before TestCondition Check End ========================

    OptCheckCounter = 0
        
    'T
    'PIN   NSIS-5 Escape Point
    'F_UNIT
    
    '光源設定 ModOrModZ1
    If OptCond.IllumMaker = NIKON Then
        With Parameter
            If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS5" Or OptCond.IllumModel = "N-SIS5KAI" Then
                Call OptMod(.Arg(1), .Arg(0))
            ElseIf OptCond.IllumModel = "N-SIS3KAI" Then
                Call OptModZ_NSIS5(.Arg(1), .Arg(0))
            End If
        End With
    End If
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'01/29 Add H.Arikawa
'光源の条件設定省略の情報を元に動作を決定する。(光源種類に応じて)
'暫定処理を入れる。

Public Sub FW_OptModOrModZ2(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 2 Then
        Err.Raise 9999, "FW_OptEscape", "The number of FW_OptEscape's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------
    
    'Call "OptIni" in case the following object "OptCond" is nothing. (OptcondオブジェクトがNothingだったらOptIniをかける)
    Call sub_CheckOptCond
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_ILLUMINATOR_MODZ2
        Call CheckBeforeTestCondition(eMode, Parameter)
        If Not IsValidTestCondition(eMode, Parameter) Then Exit Sub
        
    End If
'=========Before TestCondition Check End ========================

    OptCheckCounter = 0
        
    'Init
    'Up   NSIS-5 Escape Point
    'Down
    
    If OptCond.IllumMaker = NIKON Then
        '現在のF値、瞳距離取得
        With Parameter
            If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS5" Or OptCond.IllumModel = "N-SIS5KAI" Then
                '瞳距離の移動先へ移動
                Call OptModZ_NSIS5(.Arg(1), .Arg(0))
            ElseIf OptCond.IllumModel = "N-SIS3KAI" Then
                'F値方向の移動先へ移動
                Call OptMod(.Arg(1), .Arg(0))
            End If
        End With
    End If
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   指定時間分Waitをする
'
'パラメータ:
'    [Arg0]      In   Wait時間(s)
'
'戻り値:
'
'注意事項:
'     TheHdw.Waitで問答無用に待つ
'
Public Sub FW_DebugWait(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 1 Then
        Err.Raise 9999, "FW_DebugWait", "The number of FW_DebugWait's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_DebugWait", "FW_DebugWait's Arg0 is invalid type." & " @ " & Parameter.ConditionName
    End If
    
    Dim dblWaitTime As Double
    dblWaitTime = Parameter.Arg(0)
    
    'Waitする
    Call TheHdw.WAIT(dblWaitTime)
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   パタンの開始をおこなう(終了をまたない)
'
'パラメータ:
'    [Arg0]      In   パタン名
'    [Arg1]      In   TSBシート名
'    [Arg2]      In   実行後のウェイトタイム(省略可能 省略時Waitなし)
'
'戻り値:
'
'注意事項:IP750 or Decoder Patは、専用で設定する。
'
Public Sub PatSet(ByVal tmpPatName As String, Optional timeSetName As String = "")

    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
    Const PAT_START_LABEL As String = "START"
            
    Call StopPattern_Halt 'EeeJob関数
    Call SetTimeOut 'EeeJob関数
    
    '分割レジスタ対応ルーチン Start
    'レジスタ設定部Only(keep_alive)：PatRun
    'レジスタ設定+駆動部:PatSet
    Dim tmpPatGroupName() As String
    Dim i As Integer
    tmpPatGroupName = Split(tmpPatName, ",")
    
    PatCheckCounter = 0
    
    For i = 0 To UBound(tmpPatGroupName)
        If i < UBound(tmpPatGroupName) Then
            With TheHdw.Digital
                Call .Timing.Load(timeSetName)
                Call .Patterns.pat(tmpPatGroupName(i)).Run(PAT_START_LABEL)
            End With
            If TheExec.RunOptions.AutoAcquire = True Then
                Dim iStatus As Long
                If TheHdw.Digital.Patgen.IsRunningAnySite = True Then      'True:Still Running
                    iStatus = 0
                ElseIf TheHdw.Digital.Patgen.IsRunningAnySite = False Then 'False:haltend or keepalive
                    iStatus = 1
                End If
                
                While (iStatus <> 1)
                    If PatCheckCounter < 999 Then
                        TheHdw.TOPT.Recall
                        PatCheckCounter = PatCheckCounter + 1
                        Call WaitSet(10 * mS)
                        Exit Sub
                    End If
                Wend
            End If
        ElseIf i = UBound(tmpPatGroupName) Then
            With TheHdw.Digital
                Call .Timing.Load(timeSetName)
                Call .Patterns.pat(tmpPatGroupName(i)).Start(PAT_START_LABEL)
            End With
        End If
    Next i
    '分割レジスタ対応ルーチン End
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'内容:
'   パターンバーストを行うマクロをCallする。
'
'パラメータ:
'    [Arg0]      In   マクロ名
'
'戻り値:
'

Public Sub FW_PatSetCustomMacroA(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
    If Parameter.ArgParameterCount <> 3 And Parameter.ArgParameterCount <> 4 Then
         Err.Raise 9999, "FW_PatSetCustomMacroA", "The number of FW_PatSetCustomMacroA's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPatGroupName As String
    Dim strTsbName As String
    Dim strMacroName As String
        
    With Parameter
        strPatGroupName = .Arg(0)
        strTsbName = .Arg(1)
        strMacroName = .Arg(2)
    End With
    
    Call Application.Run(strMacroName, strPatGroupName, strTsbName)
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub



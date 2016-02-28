Attribute VB_Name = "XEeeAuto_Common"
'概要:
'   EeeAuto内で幅広く使用され、ほかでも使用されるであろう関数群
'
'目的:
'
'
'作成者:
'   2012/02/03 Ver0.1 D.Maruyama
'   2012/04/09 Ver0.2 D.Maruyama　　FW_SeparateFailSiteGndで使用するVarBank名の定義を追加
'   2012/10/19 Ver1.2 K.Tokuyoshi   以下の関数を追加
'                                   ・m_GetLimit
'                                   ・mf_GetResult
'                                   ・ngCapture_Judge_f
'                                   ・EnableFlag_False_f
'   2012/12/25 Ver1.3 H.Arikawa     "ngCap"⇒"ngCap1"に変更
'   2012/12/26 Ver1.4 H.Arikawa     "ngCap2-5"を追加
'   2013/03/15 Ver1.5 H.Arikawa     不要処理削除
'   2013/10/28 Ver1.6 H.Arikawa     条件設定省略のフラグ化

Option Explicit

Public Const JOB_KUMAMOTO_S As Long = 0
Public Const JOB_NAGASAKI_200_S As Long = 1
Public Const JOB_NAGASAKI_300_S As Long = 2

Public Const PIN_NAME_VDDSUB As String = "__VDDSUB_PIN_NAME__"
Public Const GND_SEPARATE_APMU_UB As String = "__GND_SEPARATE_APMU_UB__"
Public Const GND_SEPARATE_CUB_UB As String = "__GND_SEPARATE_CUB_UB__"
Public Const EEE_AUTO_NOUSE_STBSUB As String = "-"
Public Const EEE_AUTO_NOUSE_RELAY As String = "-"

Public Function mf_div(ByVal val1 As Double, ByVal val2 As Double, Optional ByVal errVal As Double = 0) As Double

    If val2 <> 0# Then
        mf_div = val1 / val2
    Else
        mf_div = errVal
    End If

End Function

Public Sub m_GetLimit(ByRef dblLoLimit As Double, ByRef dblHiLimit As Double)

    Dim strArgList() As String
    Dim lngArgCnt As Long
    
    Call TheExec.DataManager.GetArgumentList(strArgList, lngArgCnt)
    dblLoLimit = val(strArgList(5 * LimitSetIndex + 0))
    dblHiLimit = val(strArgList(5 * LimitSetIndex + 1))
    
End Sub

Public Function mf_GetResult(ByVal strKey As String, ByRef pResult() As Double) As Double

    On Error GoTo ErrorExit

    Call TheDcTest.GetTempResult(strKey, pResult)

    Exit Function
    
ErrorExit:
    Call TheResult.GetResult(strKey, pResult)

End Function

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'NgCapture_Test用:Start
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'内容:
'
'
'パラメータ:
'[Arg1]         In  対象TestLabel
'[Arg2]         In  対象LoLimit
'[Arg3]         In  対象HiLimit
'[Arg4]         In  対象LimitValid
'
'
Public Function ngCapture_Judge_f() As Double  '2012/11/16 175JobMakeDebug

    On Error GoTo ErrorExit

    Dim site As Long

    'パラメータの取得
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        Err.Raise 9999, "ngCapturel_Judge_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
    End If

    'Capture判定
    Dim tmpValue1() As Double
    Dim dblLoLimit As Double
    Dim dblHiLimit As Double
    Dim dblLimValid As Double

    Call mf_GetResult(ArgArr(0), tmpValue1)
    dblLoLimit = CDbl(ArgArr(1))
    dblHiLimit = CDbl(ArgArr(2))
    dblLimValid = CDbl(ArgArr(3))

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then

            Select Case dblLimValid
                Case 1
                    If tmpValue1(site) < dblLoLimit Then
                        TheExec.Flow.EnableWord("ngCap1") = True
                        TheExec.Flow.EnableWord("ngCap2") = True
                        TheExec.Flow.EnableWord("ngCap3") = True
                        TheExec.Flow.EnableWord("ngCap4") = True
                        TheExec.Flow.EnableWord("ngCap5") = True
                    End If
                Case 2
                    If tmpValue1(site) > dblHiLimit Then
                        TheExec.Flow.EnableWord("ngCap1") = True
                        TheExec.Flow.EnableWord("ngCap2") = True
                        TheExec.Flow.EnableWord("ngCap3") = True
                        TheExec.Flow.EnableWord("ngCap4") = True
                        TheExec.Flow.EnableWord("ngCap5") = True
                    End If
                Case 3
                    If tmpValue1(site) < dblLoLimit And tmpValue1(site) > dblHiLimit Then
                        TheExec.Flow.EnableWord("ngCap1") = True
                        TheExec.Flow.EnableWord("ngCap2") = True
                        TheExec.Flow.EnableWord("ngCap3") = True
                        TheExec.Flow.EnableWord("ngCap4") = True
                        TheExec.Flow.EnableWord("ngCap5") = True
                    End If
                Case Else
            End Select

        End If
    Next site


    Exit Function

ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数

End Function

Public Function EnableFlag_False_f() As Double

    TheExec.Flow.EnableWord("ngCap1") = False
    TheExec.Flow.EnableWord("ngCap2") = False
    TheExec.Flow.EnableWord("ngCap3") = False
    TheExec.Flow.EnableWord("ngCap4") = False
    TheExec.Flow.EnableWord("ngCap5") = False

End Function
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'NgCapture_Test用:End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'内容:
'   PowerDownとDisconnectを行う。
'
'パラメータ:
'    [Arg0]      In   条件名 PowerSuppluyVoltageシートでの名称
'    [Arg1]      In   シーケンス名　PowerSequenceシートでの名称
'    [Arg2]      In   実行後のウェイトタイム(省略可能 省略時Waitなし)
'    [Arg3]      In   ピン名（ピングループ名）
'    [ArgN-1]
'    [ArgN-1]    In   サイト番号(省略された場合は全サイト)
'戻り値:
'
'注意事項:
'
Public Sub PowerDownAndDisconnect()
       
    Call PowerDown4ApmuUnderShoot
    'Pinの切り離しを行う。
    Call DisconnectAllDevicePins
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

Public Sub PatRun( _
    ByVal patGroupName As String, _
    Optional ByVal startLabel As String = "START", _
    Optional timeSetName As String = "", _
    Optional categoryName As String = "", _
    Optional selectorName As String = "" _
)
    Call RunPattern(patGroupName)

End Sub


Public Sub WaitSet(ByVal waitTime As Double)
    If TheExec.RunOptions.AutoAcquire = True Then
        Call TheHdw.TOPT.WAIT(toptTimer, waitTime * 1000)
    Else
        Call TheHdw.WAIT(waitTime)
    End If
End Sub

Public Sub InitializeEeeAutoModules()

    Call InitializeDefectInformation '電圧設定マクロの初期化
    Call InitializePowerCondition '欠陥構造体の初期化
    
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call InitializeAutoConditionModify '自動TestConditionの補正
    End If

End Sub

Public Sub UnInitializeEeeAutoModules()

    Call UninitializeDefectInformation
    Call UninitializePowerCondition
    
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call UninitializeAutoConditionModify
    End If
    
End Sub



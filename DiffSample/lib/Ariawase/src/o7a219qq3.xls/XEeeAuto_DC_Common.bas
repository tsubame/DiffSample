Attribute VB_Name = "XEeeAuto_DC_Common"
'概要:
'
'
'目的:
'   DC用演算用モジュール
'
'作成者:
'   2011/12/11 Ver0.1 D.Maruyama
'   2011/12/19 Ver0.2 D.Maruyama    Arg20から呼ぶように変更
'                                   以下の関数を追加
'                                    ResultSubtract_f
'                                    ResultCalcCommonDifferential_f
'                                    ResultCalcHSImpedanceMismatch_f
'
'   2011/12/21 Ver0.3 D.Maruyama    ResultCalcLPImpedance_f関数の値の読込先を
'                                   TheDcTest→TheResultに変更
'                                   ResultCalcStpx_fを追加
'
'   2011/12/22 Ver0.4 D.Maruyama    TestInstanceからのArgの取り出しを関数化
'
'   2012/02/01 Ver0.5 D.Maruyama    DCTestScenarioのPreBody関数を追加
'
'   2012/03/07 Ver0.6 D.Maruyama    CalcOneLSB_fをジャッジするまで行うように変更
'   2012/03/17 Ver0.7 D.Maruyama    以下3つの関数を追加
'                                   ・ResultSTVD_f
'                                   ・ResultFPZR_f
'                                   ・ResultVCM_f
'   2012/10/19 Ver0.8 K.Tokuyoshi   以下の関数を追加
'                                   ・ResultCalcLPImpedance_Nega_f
'                                   ・ResultDiv_f
'                                   ・ResultMultiply_f
'                                   ・ResultDiv_Abs_f
'                                   ・ResultAbs_Sum_f
'                                   ・ResultSubtract_2_f
'                                   ・ResultCalcCommonDifferential_f
'                                   ・ResultCalcHSImpedance_f
'                                   ・ResultCalcLPImpedance_f
'                                   ・ResultCalcImpedance_f
'                                   ・ResultMin_f
'                                   ・ResultCompare_f
'                                   ・ResultSubstitution_f
'                                   ・postDcTestCommonCondition_f
'                                 　以下の関数を修正
'                                   ・ResultMax_f
'                                 　以下の関数を名前変更
'                                   ・ResultAbsDifferetial_f　⇒　ResultSubtract_Abs_f
'                                   ・ResultCalcCommonDifferential_f　⇒　ResultCalcCommonDifferential_Abs_f
'                                   ・ResultCalcHSImpedance_f　⇒　ResultCalcHSImpedance_Posi_Nega_f
'                                   ・ResultCalcLPImpedance_f　⇒　ResultCalcLPImpedance_Posi_f
'                                   ・ResultCalcStpx_f　⇒　ResultPixcel_Leak_Ratio_f
'                                   ・CalcOneLSB_f　⇒　ResultCalcOneLSB_f
'                                 　TheResult.GetResultからmf_GetResultへ変更
'   2012/10/22 Ver0.9 K.Tokuyoshi   以下の関数を追加
'                                   ・ResultBinningFlag_f
'                                   ・ResultBinning_f
'                                   ・Hold_Voltage_Test_f
'   2012/10/23 Ver1.0 K.Tokuyoshi   以下の関数を追加
'                                   ・ResultCalcStbyDifferential_f
'                                   ・ResultCalcStbyDifferential_Square_f
'                                   ・ResultCalcPiezoImpedance_f
'                                   ・ResultCalcConductance_f
'                                   ・ResultCalcConductance_Posi_Nega_f
'                                   ・ResultCalcConductance_Mono_f
'                                   ・ResultIndividualCalibrate_f
'   2013/01/31 Ver1.1 K.Hamada       以下の関数を追加
'                                   ・ResultCalcHS0_HS1Impedance_f
'   2013/02/05 Ver1.2 K.Hamada       以下の関数を修正
'                                   ・ResultCalcLPImpedance_f
'                                   ・ResultMin_f
'   2013/02/06 Ver1.3 K.Hamada       以下の関数を修正
'                                   ・ResultCalcPiezoImpedance_f
'                                   ・ResultCalcConductance_f
'                                   ・ResultCalcConductance_Posi_Nega_f
'                                   ・ResultCalcConductance_Mono_f
'   2013/02/07 Ver1.4 H.Arikawa      以下の関数を修正
'                                   ・ResultCalcLPImpedance_Posi_f (ラインデバッグFB)
'                                   ・ResultCalcPiezoImpedance_f (DC WGデバッグFB)
'                                   ・ResultCalcConductance_f    (DC WGデバッグFB)
'                                   ・ResultCalcConductance_Posi_Nega_f(DC WGデバッグFB)
'                                   ・ResultCalcConductance_Mono_f  (DC WGデバッグFB)
'                                   ・ResultCalcHS0_HS1Impedance_f  (ラインデバッグFB)
'                                   ・ResultCalcOneLsbBasic_f　　(追加)
'   2013/02/07 Ver1.5 H.Arikawa      以下の関数を修正
'                           　      ・subCurrent_Serial_NoPattern_Test_f
'　                                 ・SubCurrentTest_NoPattern_GetParameter
'   2013/02/08 Ver1.6 K.Hamada      以下の関数を修正 ※Argの指定を変更
'                                   ・ResultCalcHSImpedance_Down_f
'                                   ・ResultCalcHSImpedance_Up_f
'                                   ・ResultCalcLPImpedance_Nega_f
'   2013/02/12 Ver1.7 H.Arikawa     以下の関数を修正
'                                   ・ResultCalcOneLSB_f
'   2013/02/12 Ver1.8 H.Arikawa     以下の関数を修正
'                                   ・ResultCalcHS0_HS1Impedance_f
'   2013/02/18 Ver1.9 H.Arikawa     以下の関数を修正
'                                   ・ResultCalcOneLsbBasic_f
'   2013/02/19 Ver1.A K.Hamada      以下の関数を追加 ※Argの指定を変更
'                                   ・ResultSubtractDiv_f
'                                   ・ReturnMaxMinDiff_f
'                                   ・ReturnAbsMaxMinValueDCK_f
'   2013/02/25 Ver2.0 H.Arikawa     以下の関数を追加 ※Argの指定を変更
'                                   ・ResultMultiply_f

Option Explicit

Private Const EEE_AUTO_HOLDVOLTAGE_ARGS As Long = 9

'内容:
'   DCTestScenarioFWの共通PreBody
'
'パラメータ:
'[Arg1]         In  コンディション1
':
'[ArgN]         In　コンディションN
'
'注意事項:
'   記述されている順番にTestConditionをコールする
'   #EOPを忘れないこと
Public Function preDcTestCommonCondition_f(argc As Long, argv() As String) As Long

    On Error GoTo ErrorExit

    Call SiteCheck
    Dim i As Long
    
    If argc = 0 Then
        'エラーにしなくてよい?
        Exit Function
    End If
        
    'コンディションを順番に実施
    For i = 0 To argc - 1
        TheCondition.SetCondition argv(i)
    Next i
    
    preDcTestCommonCondition_f = TL_SUCCESS

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    preDcTestCommonCondition_f = TL_ERROR

End Function

'内容:
'   DCTestScenarioFWの共通PostBody
'
'パラメータ:
'[Arg1]         In  コンディション1
':
'[ArgN]         In　コンディションN
'
'注意事項:
'   記述されている順番にTestConditionをコールする
'   #EOPを忘れないこと
Public Function postDcTestCommonCondition_f(argc As Long, argv() As String) As Long

    On Error GoTo ErrorExit

    Call SiteCheck
    Dim i As Long
    
    If argc = 0 Then
        'エラーにしなくてよい?
        Exit Function
    End If
        
    'コンディションを順番に実施
    For i = 0 To argc - 1
        TheCondition.SetCondition argv(i)
    Next i
    
    postDcTestCommonCondition_f = TL_SUCCESS

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    postDcTestCommonCondition_f = TL_ERROR

End Function

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'IMX145用演算VBAマクロ :Start
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪


'内容:
'   TestInstanceに書かれたキーから平均をとる
'
'パラメータ:
'[Arg1]         In  対象Arg先頭
':
'[ArgN]         In　対象Arg最後
'[ArgN+1]       In  #EOP(End Of Param)
'
'注意事項:
'   SUM(Arg1,Arg2,････,ArgN) / N を計算
'   #EOPを忘れないこと
Public Function ResultAverage_f() As Double

    On Error GoTo ErrorExit

    Call SiteCheck
    
    '本関数に対するパラメータ数は不定。; The number of arguments defined on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数より小さければエラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultAverage_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのもだめ; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultAverage_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '足し合わせ
    Dim tmpValue() As Double
    For i = 0 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                retResult(site) = retResult(site) + tmpValue(site)
            End If
        Next site
        Erase tmpValue
    Next i
    
    '割戻し
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = mf_div(retResult(site), lCount)
        End If
    Next site

    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーから最大値をとる
'
'パラメータ:
'[Arg1]         In  対象Arg先頭
':
'[ArgN]         In　対象Arg最後
'[ArgN+1]       In  #EOP(End Of Param)
'
'注意事項:
'   MAX(Arg1,Arg2,････,ArgN) を計算
'   #EOPを忘れないこと
' 2012/10/19 K.Tokuyoshi Startの比較を追加
Public Function ResultMax_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '本関数に対するパラメータ数は不定。; The number of arguments defined on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultMax_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのもだめ; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultMax_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    
    'MAX算出
    Dim tmpValue0() As Double
    Dim tmpValue1() As Double
    Call mf_GetResult(ArgArr(0), tmpValue0)
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue0(site)
        End If
    Next site
    
    For i = 1 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue1)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If (retResult(site) < tmpValue1(site)) Then retResult(site) = tmpValue1(site)
            End If
        Next site
        Erase tmpValue1
    Next i
    

    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーから合計値をとる
'
'パラメータ:
'[Arg1]         In  対象Arg先頭
':
'[ArgN]         In　対象Arg最後
'[ArgN+1]       In  #EOP(End Of Param)
'
'注意事項:
'   SUM(Arg1,Arg2,････,ArgN) を計算
'   #EOPを忘れないこと
Public Function ResultSum_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '本関数に対するパラメータ数は不定。; The number of arguments defined on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得; To obtain the arguments.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultSum_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのもだめ; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultSum_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    
    '足し合わせ
    Dim tmpValue() As Double
    For i = 0 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                retResult(site) = retResult(site) + tmpValue(site)
            End If
        Next site
        Erase tmpValue
    Next i
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーから差分の絶対値をとる
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'
'注意事項:
'   Abs(Arg1-Arg2)を計算する
'
Public Function ResultSubtract_Abs_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得; To obtain the arguments.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultSubtract_Abs_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultSubtract_Abs_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultSubtract_Abs_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '絶対値差分
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = Abs(tmpValue1(site) - tmpValue2(site))
        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーから差分をとる
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'
'注意事項:
'   Arg1-Arg2を計算する
'
Public Function ResultSubtract_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得; To obtain the arguments.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultSubtract_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultSubtract_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultSubtract_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '差分
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue1(site) - tmpValue2(site)
        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーから差分の絶対値をとって2で割る
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'
'注意事項:
'   Abs(Arg1-Arg2)/2を計算する
'
Public Function ResultCalcCommonDifferential_Abs_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得; To obtain the arguments.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcCommonDifferential_Abs_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcCommonDifferential_Abs_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcCommonDifferential_Abs_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = Abs(tmpValue1(site) - tmpValue2(site)) / 2
        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーからHSインピーダンスを計算する
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'[Arg3]         In  対象Arg3
'[Arg4]         In  対象定数A
'
'注意事項:
'   (Arg1-Arg2)/Arg3/定数A　を計算する
'
Public Function ResultCalcHSImpedance_Down_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 5
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得; To obtain the arguments.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcHSImpedance_Down_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcHSImpedance_Down_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcHSImpedance_Down_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    'インピーダンス計算
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double
    Dim dblCalc1 As Double

    Call mf_GetResult(ArgArr(1), tmpValue1)
    Call mf_GetResult(ArgArr(2), tmpValue2)
    Call mf_GetResult(ArgArr(3), tmpValue3)
    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultCalcHSImpedance_Down_f", "Argument type is Mismatch """ & ArgArr(3) & """ @ " & GetInstanceName
    End If
    dblCalc1 = CDbl(ArgArr(0))

    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult(site) = mf_div(tmpValue3(site), dblCalc1)
            retResult(site) = mf_div((tmpValue2(site) - tmpValue1(site)), Temp_retResult(site))
        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

Public Function ResultCalcHSImpedance_Up_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 5
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得; To obtain the arguments.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcHSImpedance_Up_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcHSImpedance_Up_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcHSImpedance_Up_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    'インピーダンス計算
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double
    Dim dblCalc1 As Double

    Call mf_GetResult(ArgArr(1), tmpValue1)
    Call mf_GetResult(ArgArr(2), tmpValue2)
    Call mf_GetResult(ArgArr(3), tmpValue3)
    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultCalcHSImpedance_Up_f", "Argument type is Mismatch """ & ArgArr(3) & """ @ " & GetInstanceName
    End If
    dblCalc1 = CDbl(ArgArr(0))

    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult(site) = mf_div(tmpValue3(site), dblCalc1)
            retResult(site) = mf_div((tmpValue1(site) - tmpValue2(site)), Temp_retResult(site))
        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーから差分の絶対値をとって平均値でわる
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'
'注意事項:
'   Abs(Arg1-Arg2)/Ave(Arg1,Arg2)を計算する
'
Public Function ResultCalcHSImpedanceMismatch_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcHSImpedanceMismatch_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcHSImpedanceMismatch_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcHSImpedanceMismatch_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = mf_div(2 * Abs(tmpValue1(site) - tmpValue2(site)), (tmpValue1(site) + tmpValue2(site)))
        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーからLPインピーダンスを計算する
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'[Arg3]         In  対象Arg3
'[Arg4]         In  対象定数A
'
'注意事項:
'   (Arg1-Arg2)/((Arg3-Arg1)/定数A)を計算する
'
Public Function ResultCalcLPImpedance_Nega_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 5
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcLPImpedance_Nega_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcLPImpedance_Nega_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcLPImpedance_Nega_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    'インピーダンス計算
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double
    Dim dblCalc1 As Double 'Terminate R Value
    Call mf_GetResult(ArgArr(1), tmpValue1)
    Call mf_GetResult(ArgArr(2), tmpValue2)
    Call mf_GetResult(ArgArr(3), tmpValue3)
    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultCalcLPImpedance_Nega_f", "Argument type is Mismatch """ & ArgArr(3) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(0))
    
    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult(site) = mf_div((tmpValue2(site) - tmpValue3(site)), dblCalc1)
            retResult(site) = mf_div(tmpValue3(site) - tmpValue1(site), Temp_retResult(site))
        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)
    
End Function

'内容:
'   TestInstanceに書かれたキーからLPインピーダンスを計算する
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'[Arg3]         In  対象Arg3
'[Arg4]         In  対象定数A
'
'注意事項:
'   (Arg1-Arg2)/((Arg2-Arg3)/定数A)を計算する
'
Public Function ResultCalcLPImpedance_Posi_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 5
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcLPImpedance_Posi_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcLPImpedance_Posi_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcLPImpedance_Posi_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    'インピーダンス計算
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double
    Dim dblCalc1 As Double 'Terminate R Value
    Call mf_GetResult(ArgArr(1), tmpValue1)
    Call mf_GetResult(ArgArr(2), tmpValue2)
    Call mf_GetResult(ArgArr(3), tmpValue3)
    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultCalcLPImpedance_Posi_f", "Argument type is Mismatch """ & ArgArr(3) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(0))
    
    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult(site) = mf_div((tmpValue2(site) - tmpValue3(site)), dblCalc1)
            retResult(site) = mf_div(tmpValue1(site) - tmpValue2(site), Temp_retResult(site))
        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)
    
End Function

'内容:
'   TestInstanceに書かれたキーから除算を行う
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'[Arg3]         In  対象Arg3
'
'注意事項:
'  ( (Arg1+Arg2) / 2   ) / Arg3 を計算する
'
'注意事項:
'   #EOPを忘れないこと

Public Function ResultPixcel_Leak_Ratio_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 4
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数より小さければエラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultPixcel_Leak_Ratio_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultPixcel_Leak_Ratio_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultPixcel_Leak_Ratio_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '比率を算出
    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    Call mf_GetResult(ArgArr(2), tmpValue3)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult(site) = (tmpValue1(site) + tmpValue2(site)) / 2
            retResult(site) = mf_div(Temp_retResult(site), tmpValue3(site))
        End If
    Next site

    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function
'内容:
'   TestInstanceに書かれたキーからOneLsbを計算し、登録する
'
'パラメータ:
'[Arg1]         In  LSBにするための係数
'[Arg2]         In  テスト結果のキー
'
'注意事項:
'
'
Public Function ResultCalcOneLSB_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult
    
    'パラメータの取得
    '想定数より小さければエラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcOneLSB_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcOneLSB_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcOneLSB_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    'テスト結果の取得、変数ののせかえ
    Dim tmpValue() As Double
    Dim strLSBName As String
    Dim dblMultiply As Double
    Call mf_GetResult(ArgArr(1), tmpValue)
    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultCalcOneLSB_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblMultiply = CDbl(ArgArr(0))
    
    'LSBの計算
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue(site) * dblMultiply
        End If
    Next site

    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)
    
End Function
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'IMX145用演算VBAマクロ :End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'IMX145以外演算VBAマクロ :Start
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'内容:
'   TestInstanceに書かれたキーから除算を行う
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'
'注意事項:
'   (Arg1/Arg2)を計算する
'
'注意事項:
'   #EOPを忘れないこと

Public Function ResultDiv_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数より小さければエラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultDiv_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultDiv_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultDiv_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '割り算を算出
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = mf_div(tmpValue1(site), tmpValue2(site))
        End If
    Next site

    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーから除算を行う
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象定数A
'
'注意事項:
'   Arg1 * 定数A を計算する
'
'注意事項:
'   #EOPを忘れないこと

Public Function ResultMultiply_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数より小さければエラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultMultiply_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultMultiply_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultMultiply_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '掛け算を算出
    Dim tmpValue1() As Double
    Dim dblCalc1 As Double
    Call mf_GetResult(ArgArr(1), tmpValue1)
    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultMultiply_f", "Argument type is Mismatch """ & ArgArr(0) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(0))
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue1(site) * dblCalc1
        End If
    Next site

    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーから除算を行う
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象定数A
'
'注意事項:
'   |Arg1| / 定数Aを計算する
'
'注意事項:
'   #EOPを忘れないこと

Public Function ResultDiv_Abs_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数より小さければエラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultDiv_Abs_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultDiv_Abs_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultDiv_Abs_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '割り算を算出
    Dim tmpValue1() As Double
    Dim dblCalc1 As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    
    If Not IsNumeric(ArgArr(1)) Then
        Err.Raise 9999, "ResultDiv_Abs_f", "Argument type is Mismatch """ & ArgArr(1) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(1))
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = mf_div(Abs(tmpValue1(site)), dblCalc1)
        End If
    Next site

    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーから合計値をとる
'
'パラメータ:
'[Arg1]         In  対象Arg先頭
':
'[ArgN]         In　対象Arg最後
'[ArgN+1]       In  #EOP(End Of Param)
'
'注意事項:
'   Sum(Arg1,Arg2,････,ArgN) を計算
'   #EOPを忘れないこと
Public Function ResultAbs_Sum_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number of arguments on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数より小さければエラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_VARIABLE_PARAM) Then
        Err.Raise 9999, "ResultAbs_Sum_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのもだめ; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultAbs_Sum_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '合計算出
    Dim tmpValue() As Double
    For i = 0 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                retResult(site) = retResult(site) + Abs(tmpValue(site))
            End If
        Next site
        Erase tmpValue
    Next i
    

    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーから差分をとる
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'[Arg3]         In  対象Arg3
'
'注意事項:
'   Arg1-Arg2-Arg3を計算する
'
Public Function ResultSubtract_2_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 4
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultSubtract_2_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultSubtract_2_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultSubtract_2_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '3つの値差分
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double '2012/11/15 175Debug Arikawa
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    Call mf_GetResult(ArgArr(2), tmpValue3)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue1(site) - tmpValue2(site) - tmpValue3(site)
        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーから差分をとって2で割る
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'
'注意事項:
'   (Arg1-Arg2)/2を計算する
'
Public Function ResultCalcCommonDifferential_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcCommonDifferential_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcCommonDifferential_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcCommonDifferential_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = (tmpValue1(site) - tmpValue2(site)) / 2
        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーからHSインピーダンスを計算する
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象定数A
'[Arg3]         In  対象定数B
'
'注意事項:
'   定数A/(Arg1/定数B) - 定数B　を計算する
'
Public Function ResultCalcHSImpedance_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 4
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcHSImpedance_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcHSImpedance_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcHSImpedance_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    'インピーダンス計算
    Dim tmpValue1() As Double
    Dim dblCalc1 As Double
    Dim dblCalc2 As Double

    Call mf_GetResult(ArgArr(0), tmpValue1)
    
    If Not IsNumeric(ArgArr(1)) Then
        Err.Raise 9999, "ResultCalcHSImpedance_f", "Argument type is Mismatch """ & ArgArr(1) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(1))
    
    If Not IsNumeric(ArgArr(2)) Then
        Err.Raise 9999, "ResultCalcHSImpedance_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc2 = CDbl(ArgArr(2))

    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult(site) = mf_div(tmpValue1(site), dblCalc2)
            retResult(site) = mf_div(dblCalc1, Temp_retResult(site)) - dblCalc2
        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーからLPインピーダンスを計算する
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'[Arg3]         In  対象定数A
'
'注意事項:
'   |Arg1 - Arg2| / 定数A　を計算する
'
Public Function ResultCalcLPImpedance_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 4
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult
    
    'パラメータの取得
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcLPImpedance_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcLPImpedance_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcLPImpedance_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    'インピーダンス計算
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim dblCalc1 As Double

    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    If Not IsNumeric(ArgArr(2)) Then
        Err.Raise 9999, "ResultCalcLPImpedance_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(2))

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = mf_div(Abs(tmpValue1(site) - tmpValue2(site)), dblCalc1)
        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーからインピーダンスを計算する
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'[Arg3]         In  対象定数A
'
'注意事項:
'   ||Arg1| - |Arg2|| / 定数A　を計算する
'
Public Function ResultCalcImpedance_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 4
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcImpedance_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcImpedance_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcImpedance_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    'インピーダンス計算
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim dblCalc1 As Double

    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    If Not IsNumeric(ArgArr(2)) Then
        Err.Raise 9999, "ResultCalcImpedance_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(2))

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = mf_div(Abs(Abs(tmpValue1(site)) - Abs(tmpValue2(site))), dblCalc1)
        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーから最大値をとる
'
'パラメータ:
'[Arg1]         In  対象Arg先頭
':
'[ArgN]         In　対象Arg最後
'[ArgN+1]       In  #EOP(End Of Param)
'
'注意事項:
'   MIN(Arg1,Arg2,････,ArgN) を計算
'   #EOPを忘れないこと
Public Function ResultMin_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number of arguments on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数より小さければエラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultMin_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのもだめ; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultMin_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    
    'MIN算出
    Dim tmpValue() As Double
    Call mf_GetResult(ArgArr(0), tmpValue)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue(site)
        End If
    Next site
    
    For i = 1 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If (retResult(site) > tmpValue(site)) Then retResult(site) = tmpValue(site)
            End If
        Next site
        Erase tmpValue
    Next i
    

    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーから除算を行う
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'
'注意事項:
'   Arg1 < Arg2 の場合に結果に0を入れ、Arg1 > Arg2 の場合に結果に1を入れる
'
'注意事項:
'   #EOPを忘れないこと

Public Function ResultCompare_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数より小さければエラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCompare_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCompare_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCompare_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '割り算を算出
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If tmpValue1(site) < tmpValue2(site) Then
                retResult(site) = 0
            Else
                retResult(site) = 1
            End If
        End If
    Next site

    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーから値を代入する
'
'パラメータ:
'[Arg1]         In  対象Arg1
'
'注意事項:
'   代入
'   #EOPを忘れないこと
Public Function ResultSubstitution_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 2
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数より小さければエラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultSubstitution_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultSubstitution_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultSubstitution_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '代入
    Dim tmpValue1() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue1(site)
        End If
    Next site

    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーからBinning用のFlagを格納する
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2⇒定数
'
'注意事項:
'   Arg1がSpec内であれば【0】にSpec外であれば【1】にする
'   Arg2はLimitの制限範囲を記載する
'
Public Function ResultBinningFlag_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultBinningFlag_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultBinningFlag_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultBinningFlag_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '測定値/Limit範囲_Get
    Dim tmpValue1() As Double
    Dim dblbinning1 As Double
    Call mf_GetResult(ArgArr(0), retResult)
    
    If Not IsNumeric(ArgArr(1)) Then
        Err.Raise 9999, "ResultBinningFlag_f", "Argument type is Mismatch """ & ArgArr(1) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblbinning1 = CDbl(ArgArr(1))
    
    'Limit_Get
    Dim LoLimit As Double '2012/11/15 175Debug Arikawa
    Dim HiLimit As Double '2012/11/15 175Debug Arikawa
    Call m_GetLimit(LoLimit, HiLimit)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
        
                Select Case dblbinning1
                    Case 0
                            tmpValue1(site) = 0
                    Case 1
                        If retResult(site) < LoLimit Then
                            tmpValue1(site) = 1
                        End If
                    Case 2
                        If retResult(site) > HiLimit Then
                            tmpValue1(site) = 1
                        End If
                    Case 3
                        If retResult(site) < LoLimit And retResult(site) > HiLimit Then
                            tmpValue1(site) = 1
                        End If
                    Case Else
                End Select

        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)
    
    'その後のBinningの結果項目で使用できるようにResultManagerに登録しておく
    Call TheResult.Add("Flg_" & UCase(GetInstanceName), tmpValue1)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーに"Flg_"を付け結果の合計値をとる
'
'パラメータ:
'[Arg1]         In  対象Arg先頭
':
'[ArgN]         In　対象Arg最後
'[ArgN+1]       In  #EOP(End Of Param)
'
'注意事項:
'   SUM("Flg_"Arg1,"Flg_"Arg2,････,"Flg_"ArgN) を計算
'   #EOPを忘れないこと
Public Function ResultBinning_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number of arguments on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数より小さければエラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultBinning_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのもだめ; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultBinning_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    
    'Flagを足し合わせる
    Dim tmpValue() As Double
    For i = 0 To lCount - 1
        Call mf_GetResult("Flg_" & ArgArr(i), tmpValue)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                retResult(site) = retResult(site) + tmpValue(site)
            End If
        Next site
        Erase tmpValue
    Next i
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーから特殊計算①を計算し、登録する
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'[Arg3]         In  定数A
'[Arg4]         In  定数B
'
'注意事項:
'　Arg1-(A * (Arg2) + B)を計算する
'
Public Function ResultCalcStbyDifferential_f() As Double

    On Error GoTo ErrorExit
        
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 5
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult
    
    'パラメータの取得
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcStbyDifferential_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcStbyDifferential_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcStbyDifferential_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    'テスト結果の取得、変数ののせかえ
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim dblSlope As Double
    Dim dblIntercept As Double
    Call TheResult.GetResult(ArgArr(0), tmpValue1)
    Call TheResult.GetResult(ArgArr(1), tmpValue2)
    
    If Not IsNumeric(ArgArr(2)) Then
        Err.Raise 9999, "ResultCalcStbyDifferential_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblSlope = CDbl(ArgArr(2))
    
    If Not IsNumeric(ArgArr(3)) Then
        Err.Raise 9999, "ResultCalcStbyDifferential_f", "Argument type is Mismatch """ & ArgArr(3) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblIntercept = CDbl(ArgArr(3))
    
    
    'Arg2-(A * (Arg1) + B)の計算
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue1(site) - ((dblSlope * tmpValue2(site)) + dblIntercept)
        End If
    Next site
    
    'ジャッジする
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)
    
End Function

'内容:
'   TestInstanceに書かれたキーから特殊計算②を計算し、登録する
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'[Arg3]         In  定数A
'[Arg4]         In  定数B
'[Arg5]         In  定数C
'
'注意事項:
'　Arg1-(A * (Arg2)^2 + B * (Arg2) + C)を計算する
'
Public Function ResultCalcStbyDifferential_Square_f() As Double

    On Error GoTo ErrorExit
        
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 6
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult
    
    'パラメータの取得
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcStbyDifferential_Square_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If

    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcStbyDifferential_Square_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcStbyDifferential_Square_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    'テスト結果の取得、変数ののせかえ
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim dblSlope1 As Double
    Dim dblSlope2 As Double
    Dim dblIntercept As Double
    Call TheResult.GetResult(ArgArr(0), tmpValue1)
    Call TheResult.GetResult(ArgArr(1), tmpValue2)
    
    If Not IsNumeric(ArgArr(2)) Then
        Err.Raise 9999, "ResultCalcStbyDifferential_Square_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblSlope1 = CDbl(ArgArr(2))
    
    If Not IsNumeric(ArgArr(3)) Then
        Err.Raise 9999, "ResultCalcStbyDifferential_Square_f", "Argument type is Mismatch """ & ArgArr(3) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblSlope2 = CDbl(ArgArr(3))
    
    If Not IsNumeric(ArgArr(4)) Then
        Err.Raise 9999, "ResultCalcStbyDifferential_Square_f", "Argument type is Mismatch """ & ArgArr(4) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblIntercept = CDbl(ArgArr(4))
    
    
    'Arg1-(A * (Arg2)^2 + B * (Arg2) + C)の計算
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue1(site) - (dblSlope1 * (tmpValue2(site)) ^ 2 + dblSlope2 * tmpValue2(site) + dblIntercept)
        End If
    Next site
    
    'ジャッジする
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)
    
End Function
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'IMX145以外演算VBAマクロ :End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'固体値調整 :Start
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
'内容:
'   TestInstanceに書かれたキーから合計値をとり、書き込みレジスタ取得
'
'パラメータ:
'[Arg1]         In  対象Arg先頭
':
'[ArgN]         In　対象Arg最後
'[ArgN+1]       In  #EOP(End Of Param)
'
'注意事項:
'   SUM(Arg1,Arg2,････,ArgN)を計算し、[ParameterTable]シートからレジスタを取得する
'   #EOPを忘れないこと
Public Function ResultIndividualCalibrate_f() As Double

    On Error GoTo ErrorExit

    Call SiteCheck
    
    Dim site As Long
    Dim i As Long
    
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数より小さければエラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_VARIABLE_PARAM) Then
        Err.Raise 9999, "ResultIndividualCalibrate_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのもだめ
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultIndividualCalibrate_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '足し合わせ
    Dim tmpValue() As Double
    For i = 0 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                retResult(site) = retResult(site) + tmpValue(site)
            End If
        Next site
        Erase tmpValue
    Next i
    
    'シートと比較し、レジスタを取得する
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = ShtParaTable.GetREG(UCase(GetInstanceName), retResult(site))
        End If
    Next

    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'固体値調整 :End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'Hold_Voltage_Test:Start
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

Private Function Hold_Voltage_Test_f() As Double

    On Error GoTo ErrorExit

    'これひとつでひとつのテストするのでSiteCheckは必要
    Call SiteCheck
        
    '変数定義
    Dim strResultKey As String              'Arg20　項目名
    Dim strForcePin As String               'Arg21　フォース端子
    Dim strMeasurePin As String             'Arg22　テスト(測定)端子
    Dim dblStartVoltage As Double           'Arg23　Start電圧
    Dim dblEndVoltage As Double             'Arg24　End電圧
    Dim dblStepVoltage As Double            'Arg25　Step電圧
    Dim dblTargetCurrent As Double          'Arg26　Target電流
    Dim strSetParamCondition As String      'Arg27　測定パラメータ_Opt_リレー
    Dim strPowerCondition As String         'Arg28　Set_Voltage_端子設定
            
    '測定パラメータ
    Dim lAve As Double                      'TestCondition
    Dim dblClampCurrent As Double           'TestCondition
    Dim dblWait As Double                   'TestCondition
        
    '結果変数
    Dim retResult(nSite) As Double
            
    '関数内変数
    Dim dRetValue(nSite) As Double
    Dim resultI(nSite) As Double
    Dim resultID(nSite) As Double
    Dim exitflg(nSite) As Integer
    Dim exitflgJudge As Integer
    Dim TempValue(nSite) As Double
    Dim site As Long
            
    '変数取り込み
    If Not HoldVoltageTest_GetParameter( _
                strResultKey, _
                strForcePin, _
                strMeasurePin, _
                dblStartVoltage, _
                dblEndVoltage, _
                dblStepVoltage, _
                dblTargetCurrent, _
                strSetParamCondition, _
                strPowerCondition) Then
                MsgBox "The Number of Hold_Voltage_Test_f's arguments is invalid!"
                Call DisableAllTest 'EeeJob関数
                Exit Function
    End If
            
    'パラメータ設定の関数を呼ぶ (FW_SetHoldVoltageParam)
    Call TheCondition.SetCondition(strSetParamCondition)

    lAve = GetHoldVoltageAverageCount(GetInstanceName)
    dblClampCurrent = GetHoldVoltageClampCurrent(GetInstanceName)
    dblWait = GetHoldVoltageWaitTime(GetInstanceName)
    
    'High-HOLD（反転）電圧測定かLow-HOLD（反転）電圧測定で変わる部分は追加するか？？？それとも端子設定に載せてもらうか？？
    
    'HOLD電圧測定
    Dim i As Long '2012/11/15 175Debug Arikawa
    For i = dblStartVoltage To dblEndVoltage Step dblStepVoltage
        Call SetFVMI(strForcePin, i * V, dblClampCurrent)
        TheHdw.WAIT dblWait * S
        '========== MESURE IO PINS ===============================
        Call MeasureI(strMeasurePin, dRetValue(), lAve)

        For site = 0 To nSite
            If TheExec.sites.site(site).Active And exitflg(site) = 0 Then
                If Flg_Debug = 1 Then TheExec.Datalog.WriteComment strResultKey & strMeasurePin & "  " & i & " " & dRetValue(site)
                If i = dblStartVoltage Then
                    resultI(site) = dRetValue(site)
                Else
                    resultID(site) = resultI(site) - dRetValue(site)
                    If resultID(site) >= 3 * uA Then
                        retResult(site) = i - dblStepVoltage
                        exitflg(site) = 1
                    Else
                        resultI(site) = dRetValue(site)
                    End If
                End If
                exitflgJudge = exitflgJudge + exitflg(site)
            End If
        Next site
        If exitflgJudge > nSite Then Exit For
    Next i

    '測定端子の0V印加
    Call SetFVMI(strMeasurePin, 0# * V, dblClampCurrent)
               
    'ジャッジ
    Call test(retResult)

    '答えは返さずAddするのみ
    Call TheResult.Add(strResultKey, retResult)

    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

Private Function HoldVoltageTest_GetParameter( _
    ByRef strResultKey As String, _
    ByRef strForcePin As String, _
    ByRef strMeasurePin As String, _
    ByRef dblStartVoltage As Double, _
    ByRef dblEndVoltage As Double, _
    ByRef dblStepVoltage As Double, _
    ByRef dblTargetCurrent As Double, _
    ByRef strSetParamCondition As String, _
    ByRef strPowerCondition As String _
    ) As Boolean

    '変数取り込み
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_HOLDVOLTAGE_ARGS) Then
        HoldVoltageTest_GetParameter = False
        Exit Function
    End If
    
On Error GoTo ErrHandler
    strResultKey = ArgArr(0)
    strForcePin = ArgArr(1)
    strMeasurePin = ArgArr(2)
    dblStartVoltage = CDbl(ArgArr(3))
    dblEndVoltage = CDbl(ArgArr(4))
    dblStepVoltage = CDbl(ArgArr(5))
    dblTargetCurrent = CDbl(ArgArr(6))
    strSetParamCondition = ArgArr(7)
    strPowerCondition = ArgArr(8)
On Error GoTo 0

    HoldVoltageTest_GetParameter = True
    Exit Function
    
ErrHandler:

    HoldVoltageTest_GetParameter = False
    Exit Function

End Function
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'Hold_Voltage_Test:End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'Piezo電流測定(Function):Start
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'内容:
'   TestInstanceに書かれたキーからピエゾ出力インピーダンスを計算し、登録する
'
'パラメータ:
'[Arg0]         In  Arg1
'[Arg1]         In  Arg2
'[Arg2]         In  定数A
'
'注意事項:
'　定数A * (1/Arg1- 1/Arg2)を計算する
'
Public Function ResultCalcPiezoImpedance_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    Dim site As Long
    
    Dim retResult(nSite) As Double
    Erase retResult
    
    'パラメータの取得
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        Err.Raise 9999, "ResultCalcPiezoImpedance_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If

    
    'テスト結果の取得、変数ののせかえ
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim dbldiff As Double
    Call TheResult.GetResult(ArgArr(0), tmpValue1)
    Call TheResult.GetResult(ArgArr(1), tmpValue2)
        
    If Not IsNumeric(ArgArr(2)) Then
        Err.Raise 9999, "ResultCalcPiezoImpedance_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dbldiff = CDbl(ArgArr(2))
        
        
    '定数A * (1/Arg1- 1/Arg2)の計算
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = dbldiff * (mf_div(1, tmpValue1(site)) - mf_div(1, tmpValue2(site)))
        End If
    Next site
    
    'ジャッジする
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)
    
End Function
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'Piezo電流測定(Function):End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'GCS2 コンダクタンス測定(Function):Start
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'AFE次第で必要なし

'内容:
'   TestInstanceに書かれたキーから出力コンダクタンスを計算し、登録する
'
'パラメータ:
'[Arg0]         In  Arg1
'[Arg1]         In  定数A
'
'注意事項:
'　定数A * (1/Arg1)を計算する
'
Public Function ResultCalcConductance_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    Dim site As Long
    
    Dim retResult(nSite) As Double
    Erase retResult
    
    'パラメータの取得
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 3) Then
        Err.Raise 9999, "ResultCalcConductance_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If

    
    'テスト結果の取得、変数ののせかえ
    Dim tmpValue1() As Double
    Dim dblCalc1 As Double
    Call TheResult.GetResult(ArgArr(0), tmpValue1)
        
    If Not IsNumeric(ArgArr(1)) Then
        Err.Raise 9999, "ResultCalcConductance_f", "Argument type is Mismatch """ & ArgArr(1) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(1))
        
        
    '定数A * (1/Arg1)の計算
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = dblCalc1 * (mf_div(1, tmpValue1(site)))
        End If
    Next site
    
    'ジャッジする
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)
    
End Function

'内容:
'   TestInstanceに書かれたキーから出力コンダクタンス(Posi/Nega)を計算し、登録する
'
'パラメータ:
'[Arg0]         In  Arg1
'[Arg1]         In  Arg2
'[Arg2]         In  Arg3
'[Arg3]         In  定数A(電流差)
'
'注意事項:
'　| (Arg1 - Arg2) / 定数A | - Arg3を計算する
'
Public Function ResultCalcConductance_Posi_Nega_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    Dim site As Long
    
    Dim retResult(nSite) As Double
    Erase retResult
    
    'パラメータの取得
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 5) Then
        Err.Raise 9999, "ResultCalcConductance_Posi_Nega_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If

    
    'テスト結果の取得、変数ののせかえ
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double
    Dim dblCalc1 As Double
    Call TheResult.GetResult(ArgArr(0), tmpValue1)
    Call TheResult.GetResult(ArgArr(1), tmpValue2)
    Call TheResult.GetResult(ArgArr(2), tmpValue3)
    
    If Not IsNumeric(ArgArr(3)) Then
        Err.Raise 9999, "ResultCalcConductance_Posi_Nega_f", "Argument type is Mismatch """ & ArgArr(4) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(3))
        
    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult
        
    '| (Arg1 - Arg2) / 定数A | - Arg3の計算
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = Abs(mf_div((tmpValue1(site) - tmpValue2(site)), dblCalc1)) - tmpValue3(site)
        End If
    Next site
    
    'ジャッジする
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)
    
End Function

'内容:
'   TestInstanceに書かれたキーから出力コンダクタンス単調性を計算し、登録する
'
'パラメータ:
'[Arg0]         In  Arg1
'[Arg1]         In  Arg2
'[Arg2]         In  定数A
'
'注意事項:
'　(Arg1 - Arg2) / 定数Aを計算する
'
Public Function ResultCalcConductance_Mono_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    Dim site As Long
    
    Dim retResult(nSite) As Double
    Erase retResult
    
    'パラメータの取得
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        Err.Raise 9999, "ResultCalcConductance_Mono_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If

    
    'テスト結果の取得、変数ののせかえ
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim dblCalc1 As Double
    Call TheResult.GetResult(ArgArr(0), tmpValue1)
    Call TheResult.GetResult(ArgArr(1), tmpValue2)
        
    If Not IsNumeric(ArgArr(2)) Then
        Err.Raise 9999, "ResultCalcConductance_Mono_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(2))
        
    '(Arg1 - Arg2) / 定数Aの計算
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = (mf_div(tmpValue1(site) - tmpValue2(site), dblCalc1))
        End If
    Next site
    
    'ジャッジする
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)
    
End Function
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'GCS2 コンダクタンス測定(Function):End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'2013/01/31
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
'内容:
'   TestInstanceに書かれたキーからHSインピーダンスを計算する
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'[Arg3]         In  対象Arg3
'[Arg4]         In  定数A

'注意事項:
'   (Arg1-Arg2)/(Arg3/定数A　を計算する
'
Public Function ResultCalcHS0_HS1Impedance_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 5
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcHS0_HS1Impedance_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcHS0_HS1Impedance_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcHS0_HS1Impedance_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    'インピーダンス計算
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double
    
    Dim dblCalc1 As Double

    Call mf_GetResult(ArgArr(1), tmpValue1)
    Call mf_GetResult(ArgArr(2), tmpValue2)
    Call mf_GetResult(ArgArr(3), tmpValue3)
    

    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultCalcHS0_HS1Impedance_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(0))

    Dim Temp_retResult1(nSite) As Double
    Dim Temp_retResult2(nSite) As Double
    
    Erase Temp_retResult1
    Erase Temp_retResult2
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult1(site) = tmpValue1(site) - tmpValue2(site)
            Temp_retResult2(site) = mf_div(tmpValue3(site), dblCalc1)
            retResult(site) = mf_div(Abs(Temp_retResult1(site)), Temp_retResult2(site))
        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function


'内容:
'   TestInstanceに書かれたキーからHSインピーダンスを計算する
'
'パラメータ:
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'[Arg3]         In  対象Arg3
'[Arg4]         In  対象定数A
'
'注意事項:
'   (Arg1-Arg2)/Arg3/定数A　を計算する
'
Public Function ResultCalcHSImpedance_Posi_Nega_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    Dim site As Long
    
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        Err.Raise 9999, "ResultCalcHSImpedance_Posi_Nega_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
    End If
    
    'インピーダンス計算
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double
    Dim dblCalc1 As Double

    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    Call mf_GetResult(ArgArr(2), tmpValue3)
    
    If Not IsNumeric(ArgArr(3)) Then
        Err.Raise 9999, "ResultCalcHSImpedance_Posi_Nega_f", "Argument type is Mismatch """ & ArgArr(3) & """ @ " & GetInstanceName
    End If
    dblCalc1 = CDbl(ArgArr(3))

    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult(site) = mf_div(tmpValue3(site), dblCalc1)
            retResult(site) = mf_div((tmpValue1(site) - tmpValue2(site)), Temp_retResult(site))
        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'2013/01/31
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'内容:
'   TestInstanceに書かれたキーからOneLsbを計算し、登録する
'
'パラメータ:
'[Arg20]         In  テスト結果のキー
'[Arg21]         In  LSBにするための係数
'[Arg22]             #EOP
'
'注意事項:
'
'
Public Function ResultCalcOneLsbBasic_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'Arg20=DC測定項目
    'Arg21=係数(パラメータ/式)
    'Arg22=#EOP
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult
    
    'パラメータの取得
    '想定数より小さければエラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcOneLSB_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcOneLSB_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcOneLSB_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    'テスト結果の取得、変数ののせかえ
    Dim tmpValue() As Double
    Dim dblMultiply As Double
    Call mf_GetResult(ArgArr(1), tmpValue)
    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultCalcOneLSB_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblMultiply = CDbl(ArgArr(0))
    
    'LSBの計算
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue(site) * dblMultiply
        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)
    
End Function

'内容:
'   TestInstanceに書かれたキーから差分をとり、係数で割る
'
'パラメータ:
'[Arg0]         In  対象Arg0 ※係数
'[Arg1]         In  対象Arg1
'[Arg2]         In  対象Arg2
'
'注意事項:
'   (Arg1-Arg2)/Arg0を計算する
'
Public Function ResultSubtractDiv_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'パラメータ数(末尾の"#EOP"を含む); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 4
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultSubtractDiv_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのも、末尾にいないのもだめ
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultSubtractDiv_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultSubtractDiv_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim dblCalc1 As Double '係数
    Call mf_GetResult(ArgArr(1), tmpValue1)
    Call mf_GetResult(ArgArr(2), tmpValue2)

    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultSubtractDiv_f", "Argument type is Mismatch """ & ArgArr(3) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(0))
    
    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult(site) = (tmpValue1(site) - tmpValue2(site))
            retResult(site) = mf_div(Temp_retResult(site), dblCalc1)
        End If
    Next site
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーから最大値と最小値をとり差分をとる
'
'パラメータ:
'[Arg1]         In  対象Arg先頭
':
'[ArgN]         In　対象Arg最後
'[ArgN+1]       In  #EOP(End Of Param)
'
Public Function ReturnMaxMinDiff_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '本関数に対するパラメータ数は不定。; The number of arguments defined on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ReturnMaxMinDiff_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのもだめ; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ReturnMaxMinDiff_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    
    'MAX算出
    Dim tmpValue0() As Double
    Dim tmpValue1() As Double
    Dim TempMaxValue(nSite) As Double
    Dim TempMinValue(nSite) As Double
    
    Call mf_GetResult(ArgArr(0), tmpValue0)
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            TempMaxValue(site) = tmpValue0(site)
            TempMinValue(site) = tmpValue0(site)
        End If
    Next site
    
    For i = 1 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue1)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If (TempMaxValue(site) < tmpValue1(site)) Then TempMaxValue(site) = tmpValue1(site)
                If (TempMinValue(site) > tmpValue1(site)) Then TempMinValue(site) = tmpValue1(site)
            End If
        Next site
        Erase tmpValue1
    Next i
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = TempMaxValue(site) - TempMinValue(site)
        End If
    Next site

    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'内容:
'   TestInstanceに書かれたキーからDCK基準で最大値と最小値をとり絶対値の大きい値をとる
'
'パラメータ:
'[Arg1]         In  対象Arg先頭
':
'[ArgN]         In　対象Arg最後
'[ArgN+1]       In  #EOP(End Of Param)
Public Function ReturnAbsMaxMinValueDCK_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '本関数に対するパラメータ数は不定。; The number of arguments defined on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ReturnAbsMaxMinValueDCK_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのもだめ; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ReturnAbsMaxMinValueDCK_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    
    'MAX算出
    Dim tmpValue0() As Double
    Dim tmpValue1() As Double
    Dim TempMaxValue(nSite) As Double
    Dim TempMinValue(nSite) As Double
    Dim DCKValue(nSite) As Double
    
    Call mf_GetResult(ArgArr(0), tmpValue0)
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            DCKValue(site) = tmpValue0(site)
        End If
    Next site
    
    Call mf_GetResult(ArgArr(1), tmpValue1)

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            TempMaxValue(site) = tmpValue1(site) - DCKValue(site)
            TempMinValue(site) = tmpValue1(site) - DCKValue(site)
        End If
    Next site
    
    
    For i = 1 + 1 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue1)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If (TempMaxValue(site) < (tmpValue1(site) - DCKValue(site))) Then TempMaxValue(site) = tmpValue1(site) - DCKValue(site)
                If (TempMinValue(site) > (tmpValue1(site) - DCKValue(site))) Then TempMinValue(site) = tmpValue1(site) - DCKValue(site)
            End If
        Next site
        Erase tmpValue1
    Next i
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If (Abs(TempMaxValue(site)) > Abs(TempMinValue(site))) Then retResult(site) = TempMaxValue(site)
            If (Abs(TempMaxValue(site)) < Abs(TempMinValue(site))) Then retResult(site) = TempMinValue(site)
        End If
    Next site

    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function


'内容:
'   消費電力を求めるマクロ。指定される電流測定テストラベルと、対応する電圧値の値を用いる。
'
'パラメータ:
'[Arg20]        In  Arg21以後のカラムに指定される電流測定テストラベルリストと同数の「電圧値」を
'                   ","(カンマ)区切りで記載する。わざわざ電圧値を指定するのは、電流測定してる
'                   テスト条件で使用している電圧値ではない値で計算をしたいという、製品仕様側の
'                   要求があるため。
'[Arg21]...     In　電流測定をしたときのテストラベル名。
Public Function ReturnPowerConsumption_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '本関数に対するパラメータ数は不定。; The number of arguments defined on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    
    'パラメータの取得
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ReturnPowerConsumption_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのもだめ; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ReturnPowerConsumption_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '最初の値(Arg20)が電圧値。
    Dim strVddValues() As String
    strVddValues = Split(ArgArr(0), ",")
    
    'テストラベル数との数の一致を確認。
    If UBound(strVddValues) <> lCount - 2 Then
        Err.Raise 9999, "ReturnPowerConsumption_f", "The number of test labels and vdd values do not match."
    End If
    
    '浮動小数点値に変換
    Dim dblVddValues() As Double
    ReDim dblVddValues(UBound(strVddValues))
    For i = 0 To UBound(strVddValues)
        If Not IsNumeric(strVddValues(i)) Then
            Err.Raise 9999, "ReturnPowerConsumption_f", "VDD value list (Arg20) must be comma separated numeric values"
        Else
            dblVddValues(i) = CDbl(strVddValues(i))
        End If
    Next i
    
    '消費電力算出
    Dim tmpIddValue() As Double
    Dim retResult(nSite) As Double
    Erase retResult
    For i = 0 To UBound(dblVddValues)
        Call mf_GetResult(ArgArr(i + 1), tmpIddValue)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                retResult(site) = retResult(site) + tmpIddValue(site) * dblVddValues(i)
            End If
        Next site
    Next i
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'大判_ScrnWait_VBAマクロ :Start
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'内容:
'   スクリーニングを実行し印加時間をテスト項目へ返す。
'
'パラメータ:
'
'注意事項:
'

Public Function ScreeninApplyWait_f() As Double

    On Error GoTo ErrorExit

    '変数定義
    'Arg 20: 印加条件
    'Arg 21: 測定要求仕様書に指定されるスクリーニングのWait時間
    'Arg 22: テストラベル
    Dim strSetCondition As String       'Arg20: [Test Condition]'s condition name excluding screening wait.
    Dim dblWaitTime As Double           'Arg21　印加電圧
    Dim strResultKey As String          'Arg22　項目名
    
    If Flg_Scrn = 1 And Flg_Tenken = 0 Then

        '変数取り込み
        If Not ScreeningWait_GetParameter( _
                    strSetCondition, _
                    dblWaitTime, _
                    strResultKey) Then
                    
                    MsgBox "The Number of ScreeninApplyWait_f's arguments is invalid!"
                    Call DisableAllTest 'EeeJob関数
                    Exit Function
                    
        End If
            
        '========== Set Condition ===============================
        Call TheCondition.SetCondition(strSetCondition)
        
        TheHdw.WAIT (dblWaitTime)    '印加時間
    Else
        Exit Function
    End If
        
    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数

End Function

Private Function ScreeningWait_GetParameter(ByRef strSetCondition As String, _
                                            ByRef dblWaitTime As Double, _
                                            ByRef strResultKey As String) As Boolean

    '変数取り込み
    '想定数よりと違う場合エラーコード
    
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        ScreeningWait_GetParameter = False
        Exit Function
    End If
    
    Dim tempstr As String      '一時保存変数
    Dim tempArrstr() As String '一時保存配列
    
On Error GoTo ErrHandler
    strSetCondition = ArgArr(0)        'Arg24-1: [Test Condition]'s condition name for common environment setup.
    dblWaitTime = ArgArr(1)       'Arg22: Force voltage (PPS value) for test pin.
    strResultKey = ArgArr(2)                'Arg20: Test label name.
On Error GoTo 0

    ScreeningWait_GetParameter = True
    Exit Function
    
ErrHandler:

    ScreeningWait_GetParameter = False
    Exit Function

End Function


'内容:
'   TestInstanceに書かれたキーと、各キーに対する重み付け係数から係数をかけた値の総和をとる。
'       (To obtain the sum of previous test results multiplied with user factors)
'メモ：
'   IMX227でのMIPIの実動作時の想定消費電流測定のために開発。MIPIの実動作では、1H期間に対する
'   時間比として、にLPモード駆動が65%、HSモード駆動が35%を占める。それぞれの動作時の消費電流量を
'   先に測定しておき、その値に0.35ならびに0.65をかけて和をとることを目的とした。
'     このような計算処理をより汎用化したものが本関数である。
'       (Originally it is intended to calculate reliable current consumption under MIPI burst mode.
'        In the test, both current consumption value under MIPI-HS and LP burst is measured
'        respectively, and then total cunsumption value is calculated by the following equation
'               [Total current] = 0.35 * [current under HS-Burst] + 0.65 * [current under LP-burst]
'       where, 0.35 and 0.65 is the ratio of the burst period in 1H
'
'
'パラメータ:
'[Arg20]        In  Arg21から"#EOP"コードのあるArgumentの-1列までに並んでいる
'                   キー(テストラベル名)と同数の「重み付け係数」が、","
'                   (半角カンマ)区切りで列挙されたもの。測定要求仕様書の
'                   「式/パラメータ」における"$係数"式の右辺。
'                       例) 0.35,0.65
'                   (Comma separated weight factor values. The number of values must equal to
'                   the number of test label names specified at [ArgN].)
'[ArgN]         In　(Nは21以上の整数)対象キー(テストラベル名)
'                   (Target test label names.)
'[ArgN+1]       In  #EOP(End Of Param)
'
'注意事項:
'   #EOPを忘れないこと
Public Function ResultWeightFactorSum_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '本関数に対するパラメータ数は不定。; The number of arguments defined on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    'パラメータの取得; To obtain the arguments.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultSum_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '終了文字列が見つからないのもだめ; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0始まりなので#EOPの位置が有効引数の数となる
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultWeightFactorSum_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '第一引数の重み付け係数の分解
    Dim strWeightFactors() As String
    strWeightFactors = Split(ArgArr(0), ",")
    If UBound(strWeightFactors) <> lCount - 2 Then
        Call MsgBox("Error occurred! : The number of test keys and factors do not match.")
        Call DisableAllTest
        Call test(retResult)
        Exit Function
    End If
    Dim dblWeightFactors() As Double
    ReDim dblWeightFactors(UBound(strWeightFactors))
On Error GoTo NotNumericError
    For i = 0 To UBound(strWeightFactors)
        dblWeightFactors(i) = CDbl(strWeightFactors(i))
    Next i
On Error GoTo ErrorExit
    
    
    '足し合わせ
    Dim tmpValue() As Double
    For i = 1 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                retResult(site) = retResult(site) + tmpValue(site) * dblWeightFactors(i - 1)
            End If
        Next site
        Erase tmpValue
    Next i
    
    'ジャッジ
    Call test(retResult)
    
    'その後のテストで使用できるようにResultManagerに登録しておく
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
NotNumericError:
    MsgBox "Error Occurred !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description & vbCrLf & "Weight Factors (Arg20 of Test Instances sheet) must be comma separated numeric values."
    Call DisableAllTest
    Call test(retResult)
    Exit Function
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)

End Function

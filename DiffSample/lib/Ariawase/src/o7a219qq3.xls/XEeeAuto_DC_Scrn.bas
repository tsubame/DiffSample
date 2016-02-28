Attribute VB_Name = "XEeeAuto_DC_Scrn"
'概要:
'
'
'目的:
'   高電圧SCRNを行うためのモジュール
'
'作成者:
'   2012/01/23 Ver0.1 D.Maruyama
'   2012/02/14 Ver0.2 D.Maruyama　SV125の値をSCRNフラグによらずAddするように変更
'   2012/03/07 Ver0.3 D.Maruyama　TestInstanceからForceTime後のパターン設定を取るように変更
'                                 SetMV後のWaitをTestConditionからもらうように変更
'   2012/10/19 Ver0.4 K.Tokuyoshi 大幅に修正
'   2012/10/26 Ver0.5 K.Tokuyoshi 以下の関数を追加
'                                 ・ResultScrnSpec_f
'   2012/11/14・11/15 Ver0.6 T.Morimoto  以下の関数を追加・修正
'                                 ・FW_DcScreeningSet、Screening_GetParameter、FW_DcScreeningStop

Option Explicit

'+
' Name      : ScreeningFlag
' Purpose   : [J]   端子電圧測定のない、単純な高電圧スクリーニング印加テストで、印加したかしなかったかを、
'                   "Flg_Scrn"の値で返す。
'             [E]   Test implementation function for simple high-voltage screening. Returns the
'                   "Flg_Scrn" value indicating screening-on/off.
' Arguments : [J]   Test InstancesシートのArg20/21/22より入力。
'             [E]   Arguments must be specified at cells Arg20,21,22 on Test Instances worksheet.
'                   Arg20   "Condition Name" for relay, illuminator, power supply, pin electronics
'                           settings, which are defined on TestCondition worksheet.
'                   Arg21   Wait time for which high voltage screening is applied.
'                   Arg22   The test label.
' Restrictions  [J] Test Instancesシートからの呼び出し限定。
'               [E] Must be called from TheExec.Flow.
' History       First drafted by TM 2014-Feb-03
'                   - For Kumamoto IMX219 Shinraisei analysis (Koyama-san).
'                   - For masterbook 013/016
'-
Private Function ScreeningFlag() As Double

    On Error GoTo ErrorExit

    Call SiteCheck
    
    '変数定義
    'Arg 20: 測定パラメータ_Opt_リレー & Set_Voltage_端子設定 & Pattern & Wait
    'Arg 21: 測定要求仕様書に指定されるスクリーニングのWait時間
    'Arg 22: スクリーニングと一緒に行われるDC測定のWait時間
    Dim strSetCondition As String       'Arg20: [Test Condition]'s condition name excluding screening wait.
    Dim dblScreeningWait As Double      'Arg21: The wait time for screening specified on the specification sheet.
    Dim testLabelName As String             'Arg22: The test label name of the test.
    
    Dim site As Long
    Dim tmpResult(nSite) As Double

    If Flg_Scrn = 0 Then
        TheResult.Add "IDDBI_HSN", tmpResult
    End If

    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmpResult(site) = Flg_Scrn And (Flg_Tenken = 0)
        End If
    Next site

    '変数取り込み
    If Not Screening_GetParameterFlag( _
                strSetCondition, _
                dblScreeningWait, _
                testLabelName) Then
                MsgBox "The Number of ScreeningFlag's arguments is invalid!"
                Call DisableAllTest 'EeeJob関数
                Exit Function
    End If

    If Flg_Scrn = 1 And Flg_Tenken = 0 Then

        '========== Set Condition ===============================
        Call TheCondition.SetCondition(strSetCondition)
        Call TheHdw.WAIT(dblScreeningWait)
        Call TheResult.Add(testLabelName, tmpResult)
    Else
        Call TheResult.Add(testLabelName, tmpResult)
        Exit Function
    End If
    
    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数

End Function

Private Function Screening_GetParameterFlag( _
    ByRef strSetCondition As String, _
    ByRef dblScreeningWait As Double, _
    ByRef testLabelName As String _
    ) As Boolean
    
    '変数取り込み
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        Screening_GetParameterFlag = False
        Exit Function
    End If
    
On Error GoTo ErrHandler
    strSetCondition = ArgArr(0)
    dblScreeningWait = ArgArr(1)
    testLabelName = ArgArr(2)
On Error GoTo 0

    Screening_GetParameterFlag = True
    Exit Function
    
ErrHandler:

    Screening_GetParameterFlag = False
    Exit Function

End Function


'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'IMX145_Scrn_VBAマクロ :Start
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'内容:
'   Flgの値を格納する。
'
'パラメータ:
'
'注意事項:
'
'
Public Function ResultScrnFlg_f() As Double

    On Error GoTo ErrorExit
        
    Call SiteCheck
    
    Dim site As Long
    
    Dim retResult(nSite) As Double
    Erase retResult
    
    If Flg_Scrn = 1 And Flg_Tenken = 0 Then
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                retResult(site) = Flg_Scrn
            End If
        Next site
    Else
    End If
    
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
'   Waitの値を格納する。
'
'パラメータ:
'
'注意事項:
'
'
Public Function ResultScrnWait_f() As Double

    On Error GoTo ErrorExit

    Call SiteCheck

    Dim site As Long

    Dim retResult(nSite) As Double
    Erase retResult

    If Flg_Scrn = 1 And Flg_Tenken = 0 Then
        'パラメータの取得
        '想定数より小さければエラーコード
        Dim ArgArr() As String
        If Not EeeAutoGetArgument(ArgArr, 1) Then
            Err.Raise 9999, "ResultScrnWait_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        End If

        'Wait時間取得
        Dim tmpValue1 As Double
        tmpValue1 = ArgArr(0) '2012/11/15 175Debug Arikawa CDbl Delete

        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                retResult(site) = tmpValue1
            End If
        Next site
    Else
    End If

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

Private Function FW_DcScreeningSet() As Double

    On Error GoTo ErrorExit

    Call SiteCheck
    
    '変数定義
    'Arg 20: 測定パラメータ_Opt_リレー & Set_Voltage_端子設定 & Pattern & Wait
    'Arg 21: 測定要求仕様書に指定されるスクリーニングのWait時間
    'Arg 22: スクリーニングと一緒に行われるDC測定のWait時間
    Dim strSetCondition As String       'Arg20: [Test Condition]'s condition name excluding screening wait.
    Dim dblScreeningWait As Double      'Arg21: The wait time for screening specified on the specification sheet.
    Dim dblMeasurementWait As Double    'Arg22: The wait time for DC measurement (V125, VBGR... etc).
    
    If Flg_Scrn = 1 And Flg_Tenken = 0 Then

        '変数取り込み
        If Not Screening_GetParameter( _
                    strSetCondition, _
                    dblScreeningWait, _
                    dblMeasurementWait) Then
                    MsgBox "The Number of FW_DcSetScreening's arguments is invalid!"
                    Call DisableAllTest 'EeeJob関数
                    Exit Function
        End If
            
        '========== Set Condition ===============================
        Call TheCondition.SetCondition(strSetCondition)
        If dblScreeningWait > dblMeasurementWait Then Call TheHdw.WAIT(dblScreeningWait - dblMeasurementWait)
    Else
        Exit Function
    End If
        

    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数

End Function

Private Function Screening_GetParameter( _
    ByRef strSetCondition As String, _
    ByRef dblScreeningWait As Double, _
    ByRef dblMeasurementWait As Double _
    ) As Boolean
    
    '変数取り込み
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        Screening_GetParameter = False
        Exit Function
    End If
    
On Error GoTo ErrHandler
    strSetCondition = ArgArr(0)
    dblScreeningWait = ArgArr(1)
    dblMeasurementWait = ArgArr(2)
On Error GoTo 0

    Screening_GetParameter = True
    Exit Function
    
ErrHandler:

    Screening_GetParameter = False
    Exit Function

End Function

Private Function FW_DcScreeningMeasure() As Double

    On Error GoTo ErrorExit

    'これひとつでひとつのテストするのでSiteCheckは必要
    Call SiteCheck
        
    '定数定義
    Const PARAM_DELIMITER As String = ","
    
    '変数定義
    Dim strSetDCMeasure As String       'Arg20　DCシナリオ
    Dim strTestLabelNames As String     'Arg21 DC Measureを行った結果のラベル名(複数カンマの可能性あり)
    Dim strDummyTestResult As String    'Arg22 DC Measureを行わなかった場合のダミーの値(複数カンマの可能性あり)
    Dim strTestLabels() As String
    Dim strDummyValues() As String
    Dim dblDummyValues(nSite) As Double
    Dim i As Long
    Dim site As Long
    
    '変数取り込み
    If Not Screening_GetMeasure( _
                strSetDCMeasure, _
                strTestLabelNames, _
                strDummyTestResult) Then
                MsgBox "The Number of FW_DcScreeningMeasure's arguments is invalid!"
                Call DisableAllTest 'EeeJob関数
                Exit Function
    End If
    
    If Flg_Scrn = 1 And Flg_Tenken = 0 Then
                        
        '========== DCシナリオシート実行 ===============================
        Call TheDcTest.SetScenario(strSetDCMeasure)
        TheDcTest.Execute
        
    Else
        'テストラベル分解
        strTestLabels = Split(strTestLabelNames, PARAM_DELIMITER)
        'ダミー測定値分解
        strDummyValues = Split(strDummyTestResult, PARAM_DELIMITER)
        '--Error処理：テストラベル数と、ダミー用の特性値が不一致の場合
        If UBound(strTestLabels) <> UBound(strDummyValues) Then
            Call MsgBox("The number of test labels and dummy values do not match. Check <parameter/equation> column on your specification sheet.")
            GoTo ErrorExit
        End If
        
        For i = 0 To UBound(strDummyValues)
            If IsNumeric(strDummyValues(i)) Then
                For site = 0 To nSite
                    dblDummyValues(site) = CDbl(strDummyValues(i))
                Next site
                Call TheResult.Add(strTestLabels(i), dblDummyValues)
            Else
                Call MsgBox("The dummy return value must be numeric.")
                GoTo ErrorExit
            End If
        Next i
        Exit Function
    End If

    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数

End Function


Private Function FW_DcScreeningStop() As Double

    On Error GoTo ErrorExit

    'これひとつでひとつのテストするのでSiteCheckは必要
    Call SiteCheck
    
    If Flg_Scrn = 1 And Flg_Tenken = 0 Then
        Call PowerDown4ApmuUnderShoot
    Else
        Exit Function
    End If
        
    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数

End Function

Private Function Screening_GetMeasure( _
    ByRef strSetDCMeasure As String, _
    ByRef strTestLabelNames As String, _
    ByRef strDummyTestResult As String _
    ) As Boolean

    '変数取り込み
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        Screening_GetMeasure = False
        Exit Function
    End If
    
On Error GoTo ErrHandler
    strSetDCMeasure = ArgArr(0)
    strTestLabelNames = ArgArr(1)
    strDummyTestResult = ArgArr(2)
On Error GoTo 0

    Screening_GetMeasure = True
    Exit Function
    
ErrHandler:

    Screening_GetMeasure = False
    Exit Function

End Function

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'SCR TOPT用　FW_SetConditionMacro:
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
'

'内容:
'   SetConditionを行う
'
'パラメータ:
'    [Arg0]      In Condition Name
'
'戻り値:
'
'注意事項:
'
Public Sub FW_SetDcScreening_topt(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    '========= TestCondition Call ======================
    If Flg_Scrn = 1 And Flg_Tenken = 0 Then
    
        If Parameter.ArgParameterCount() <> 1 Then
            Err.Raise 9999, "FW_SetDcScreening_topt", "The number of FW_SetDcScreening_topt's arguments is invalid." & " @ " & Parameter.ConditionName
        End If
                        
        '========== Set Condition ===============================
        Call TheCondition.SetCondition(Parameter.Arg(0))
        
    Else
        Exit Sub
    End If
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'SCR TOPT用　FW_MeasureMacro:
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
'

'内容:
'   Measureを行う
'
'パラメータ:
'    [Arg0]      In DC Test Scenario Name
'
'戻り値:
'
'注意事項:
'
Public Sub FW_DcMeasure_topt(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 1 Then
        Err.Raise 9999, "FW_DcMeasure_topt", "The number of FW_DcMeasure_topt's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '========= TestCondition Call ======================
    If Flg_Scrn = 1 And Flg_Tenken = 0 Then
                        
        '========== DCシナリオシート実行 ===============================
        TheDcTest.SetScenario (Parameter.Arg(0))
        TheDcTest.Execute
        
    Else
        Exit Sub
    End If
    '========= TestCondition Call ======================
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    
End Sub



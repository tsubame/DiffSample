Attribute VB_Name = "XEeeAuto_DC_SubCurrent"
'概要:
'
'
'目的:
'   STB電流測定を行うためのモジュール
'
'作成者:
'   2011/12/06 Ver0.1 D.Maruyama
'   2011/12/06 Ver0.2 D.Maruyama ピンリソース名を追加
'   2011/12/16 Ver0.3 D.Maruyama TestInstanceのArg開始位置を変更
'   2011/12/22 Ver0.4 D.Maruyama TestInstanceからのArgの取り出しを関数化
'   2012/01/23 Ver0.5 D.Maruyama パラメータの一部をTestConditionシート経由に変更
'   2012/02/03 Ver0.6 D.Maruyama Key名をTestInstance経由で取得するようにに変更
'   2012/02/14 Ver0.7 D.Maruyama SUB電流測定のパラメータ設定をこのインスタンスの
'                                TestConditionシートから呼べるように変更。
'   2012/02/20 Ver0.8 D.Maruyama ジャッジは別にしないといけないので、ResultManagerにAddするキー名を別に渡す
'   2012/03/16 Ver0.9 D.Maruyama APMUUBの切り離しがCUBになっていたのを修正
'   2012/10/19 Ver1.2 K.Tokuyoshi 大幅に修正
'   2012/12/26 Ver1.3 H.Arikawa  GndSeparateBySiteをPrivate関数からPublic関数へ変更
'   2013/01/22 Ver1.4 H.Arikawa  subCurrent_Test_fをSubCurrentTestIfNeeded_fへ変更、パラメータ設定修正。
'   2013/01/25 Ver1.5 H.Arikawa  修正。
'   2013/01/31 Ver1.6 H.Arikawa  修正。
'   2013/02/05 Ver1.7 H.Arikawa  Debug内容反映。
'   2013/02/07 Ver1.8 H.Arikawa  subCurrent_Serial_Test_fを修正。
'   2013/02/07 Ver1.9 H.Arikawa  SubCurrentTest_NoPattern_GetParameter、subCurrent_Serial_NoPattern_Test_fを追加。
'   2013/02/12 Ver2.0 H.Arikawa  Arg数定義部分修正。subCurrent_Serial_NoPattern_Test_f修正。
'   2013/02/22 Ver2.1 H.Arikawa  subCurrent_Serial_NoPattern_Test_f修正。
'   2013/03/11 Ver2.2 K.Hamada   SubCurrentTestIfNeeded_f 修正
'                                SubCurrentNonScenario_Measure_f追加
'                                subCurrentNonScenarioSeriParaJudge_f追加


Option Explicit

'パラレル測定後シリアル測定を行なう際のシリアル測定用のパラメータ数。末尾の"#EOP"も含む。
'   Number of arguments for serial current measurement following parallel measurement and its judge.
'   "#EOP" at the end is also accounted.
Public Const EEE_AUTO_SERIPARA_SERI_ARGS As Long = 4
Public Const EEE_AUTO_BPMU_PARA_ARGS As Long = 4

'シリアル測定のみの電流測定用のパラメータ数。末尾の"#EOP"も含む。
'   Number of arguments for serial current measurement only test.
'   "#EOP" at the end is also accounted.
Public Const EEE_AUTO_SERIAL_ARGS As Long = 7

'パラレル測定後シリアル測定に突入するかどうかの判定用のパラメータ数。末尾の"#EOP"も含む
'   Number of arguments for judgement of execution for serial current measurement including "#EOP" at the end.
Public Const EEE_AUTO_SUB_SERIPARA_JUDGE_ARGS As Long = 5

'シリアル測定のみの電流測定用のパラメータ数。末尾の"#EOP"も含む。(パターン無し用)
Public Const EEE_AUTO_SUBCURRENT_ARGS As Long = 6

Public Go_Serial_Mesure As Boolean

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Serial_Test用:Start
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

Private Function subCurrent_Serial_Test_f() As Double

    On Error GoTo ErrorExit

    'これひとつでひとつのテストするのでSiteCheckは必要
    Call SiteCheck
        
    '変数定義
    Dim strResultKey As String              'Arg20　項目名
    Dim strPin As String                    'Arg21　テスト端子
    Dim dblForceVoltage As Double           'Arg22　印加電圧
    Dim strSetParamCondition As String      'Arg23　測定パラメータ_Opt_リレー
    Dim strPowerCondition As String         'Arg24　Set_Voltage_端子設定
    Dim strPatternCondition As String       'Arg25　Pattern
            
    '測定パラメータ
    Dim lAve As Double                      'TestCondition
    Dim dblClampCurrent As Double           'TestCondition
    Dim dblWait As Double                   'TestCondition
    Dim dblPinResourceName As String        'TestCondition
        
    '結果変数
    Dim retResult(nSite) As Double
            
    '関数内変数
    Dim Flg_Active(nSite) As Long
    Dim TempValue(nSite) As Double
    Dim site As Long
            
    '変数取り込み
    If Not SubCurrentTest_GetParameter( _
                strResultKey, _
                strPin, _
                dblForceVoltage, _
                strSetParamCondition, _
                strPowerCondition, _
                strPatternCondition) Then
                MsgBox "The Number of subCurrent_Serial_Test_f's arguments is invalid!"
                Call DisableAllTest 'EeeJob関数
                Exit Function
    End If
            
    'パラメータ設定の関数を呼ぶ (FW_SetSubCurrentParam)
    Call TheCondition.SetCondition(strSetParamCondition)

    lAve = GetSubCurrentAverageCount(GetInstanceName)
    dblClampCurrent = GetSubCurrentClampCurrent(GetInstanceName)
    dblWait = GetSubCurrentWaitTime(GetInstanceName)
    dblPinResourceName = GetSubCurrentPinResourceName(GetInstanceName)

    'Activeサイトの確認
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Flg_Active(site) = 1
        End If
    Next site

    'SUB電流測定の確認
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            '========== 未使用siteのBetaGND切り離し ===============================
            Call GndSeparateBySite(site)
            
            '========== Set Condition ===============================
            Call TheCondition.SetCondition(strPowerCondition)

             '========== Force Voltage ===============================
            If StrComp(dblPinResourceName, "BPMU", vbTextCompare) = 0 Then
                Call SetFVMI_BPMU(strPin, 0 * V, dblClampCurrent, site)
                Call DisconnectPins(strPin, site)
                Call SetFVMI_BPMU(strPin, dblForceVoltage, dblClampCurrent, site)
            Else
                Call SetFVMI(strPin, 0 * V, dblClampCurrent, site)
                Call DisconnectPins(strPin, site)
                Call SetFVMI(strPin, dblForceVoltage, dblClampCurrent, site)
            End If
            
             '========== Set Pattern ===============================
            Call TheCondition.SetCondition(strPatternCondition)
            TheHdw.WAIT dblWait
            
             '========== Measure Current ===============================
            If StrComp(dblPinResourceName, "BPMU", vbTextCompare) = 0 Then
                Call MeasureI_BPMU(strPin, retResult, lAve, site)
                Call SetFVMI_BPMU(strPin, 0 * V, dblClampCurrent, site)
                Call DisconnectPins(strPin, site)
            Else
                Call MeasureI(strPin, retResult, lAve, site)
                Call SetFVMI(strPin, 0 * V, dblClampCurrent, site)
                Call DisconnectPins(strPin, site)
            End If

             '========== 未使用siteのBetaGND戻し =======================
            Call GndConectBySite(site, Flg_Active)
        End If
    Next site
    
    'パターン停止
    Call StopPattern 'EeeJob関数
  
    'All_Open及びDisconnect自動生成化
    Call PowerDownAndDisconnect
           
    '答えは返さずAddするのみ
    Call test(retResult)
    Call TheResult.Add(strResultKey, retResult)

    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)
    Call TheResult.Add(strResultKey, retResult)

End Function
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Serial_Test用:End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
'2013/01/22 H.Arikawa Arg23,24,25 Get -> Arg24 カンマ区切り対応
Private Function SubCurrentTest_GetParameter( _
    ByRef strResultKey As String, _
    ByRef strPin As String, _
    ByRef dblForceVoltage As Double, _
    ByRef strSetParamCondition As String, _
    ByRef strPowerCondition As String, _
    ByRef strPatternCondition As String _
    ) As Boolean

    '変数取り込み
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_SERIAL_ARGS) Then
        SubCurrentTest_GetParameter = False
        Exit Function
    End If
    
On Error GoTo ErrHandler
    strResultKey = ArgArr(0)                'Arg20: Test label name.
    strPin = ArgArr(1)                      'Arg21: Test pin name
    dblForceVoltage = CDbl(ArgArr(2))       'Arg22: Force voltage (PPS value) for test pin.
    strSetParamCondition = ArgArr(3)        'Arg23: [Test Condition]'s condition name for common environment setup.
    strPowerCondition = ArgArr(4)           'Arg24: [Test Condition]'s condition name for device setup.
    strPatternCondition = ArgArr(5)         'Arg25: [Test Condition]'s condition name for patttern burst.
On Error GoTo 0

    SubCurrentTest_GetParameter = True
    Exit Function
    
ErrHandler:

    SubCurrentTest_GetParameter = False
    Exit Function

End Function
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Serial_Test用:End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Serial_Test用:End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
Private Function getParam_SerialMeasureAfterParallel( _
    ByRef strPin As String, _
    ByRef dblForceVoltage As Double, _
    ByRef strSetParamCondition As String, _
    ByRef strPowerCondition As String, _
    ByRef strPatternCondition As String _
    ) As Boolean

    '変数取り込み
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_SERIPARA_SERI_ARGS) Then
        getParam_SerialMeasureAfterParallel = False
        Exit Function
    End If
    
    Dim tempstr As String      '一時保存変数
    Dim tempArrstr() As String '一時保存配列
    
On Error GoTo ErrHandler
    strPin = ArgArr(0)                      'Arg20: Test pin name
    dblForceVoltage = CDbl(ArgArr(1))       'Arg21: Force voltage (PPS value) for test pin.
    tempstr = ArgArr(2)
    tempArrstr = Split(tempstr, ",")
    strSetParamCondition = tempArrstr(0)        'Arg22-1: [Test Condition]'s condition name for common environment setup.
    strPowerCondition = tempArrstr(1)           'Arg22-2: [Test Condition]'s condition name for device setup.
    strPatternCondition = tempArrstr(2)         'Arg22-3: [Test Condition]'s condition name for patttern burst.
On Error GoTo 0

    getParam_SerialMeasureAfterParallel = True
    Exit Function
    
ErrHandler:

    getParam_SerialMeasureAfterParallel = False
    Exit Function

End Function


'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Serial_Test用:End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
Private Function getParam_SerialMeasureAfterParallel_NoPattern( _
    ByRef strPin As String, _
    ByRef dblForceVoltage As Double, _
    ByRef strSetParamCondition As String, _
    ByRef strPowerCondition As String _
    ) As Boolean

    '変数取り込み
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_SERIPARA_SERI_ARGS) Then
        getParam_SerialMeasureAfterParallel_NoPattern = False
        Exit Function
    End If
    
    Dim tempstr As String      '一時保存変数
    Dim tempArrstr() As String '一時保存配列
    
On Error GoTo ErrHandler
    strPin = ArgArr(0)                      'Arg20: Test pin name
    dblForceVoltage = CDbl(ArgArr(1))       'Arg21: Force voltage (PPS value) for test pin.
    tempstr = ArgArr(2)
    tempArrstr = Split(tempstr, ",")
    strSetParamCondition = tempArrstr(0)        'Arg22-1: [Test Condition]'s condition name for common environment setup.
    strPowerCondition = tempArrstr(1)           'Arg22-2: [Test Condition]'s condition name for device setup.
On Error GoTo 0

    getParam_SerialMeasureAfterParallel_NoPattern = True
    Exit Function
    
ErrHandler:

    getParam_SerialMeasureAfterParallel_NoPattern = False
    Exit Function

End Function
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Serial_Test用:End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
Private Sub GndSeparateBySite(ByVal targetSite As Long)

    Dim site As Long

    '========== 全SITE強制ACTIVE化 (実行SITE以外) =============================
    For site = 0 To nSite
        If site <> targetSite Then TheExec.sites.site(site).Active = True
    Next site
    
    '========== 分離前のデバイス停止設定 ======================================
    'パターン停止
    Call StopPattern 'EeeJob関数
          
    'All_Open及びDisconnect自動生成化
    Call PowerDownAndDisconnect
    
    '========== 未使用SITEのGND分離 ===========================================
    For site = 0 To nSite
        If site <> targetSite Then Call SET_RELAY_CONDITION("GND_Separate_Site" & CStr(site), "-") '2012/11/16 175Debug Arikawa
    Next site
                  
    '========== 未使用SITEの停止 ==============================================
    For site = 0 To nSite
        If site <> targetSite Then TheExec.sites.site(site).Active = False
    Next site

End Sub

Private Sub GndConectBySite(ByVal targetSite As Long, ByRef ActiveSiteFlg() As Long)

    Dim site As Long

    '========== 未測定SITEのGND =============================
    For site = 0 To nSite
        If site <> targetSite Then
            If ActiveSiteFlg(site) = 1 Then
                TheExec.sites.site(site).Active = True
                Call SET_RELAY_CONDITION("GND_Beta_Site" & CStr(site), "-") '2012/11/16 175Debug Arikawa
            End If
        End If
    Next site

End Sub

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Parallel→Serial_Judge用:Start
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

Private Function subCurrentSeriParaJudge_f() As Double

    'これひとつでひとつのテストするのでSiteCheckは必要
    Call SiteCheck
    
    Go_Serial_Mesure = False
    
    '変数定義
    Dim strResultKey As String              'Arg20　Test label
    Dim dblLoJudgeLimit As Double           'Arg21　Serial突入Lowリミット
    Dim dblHiJudgeLimit As Double           'Arg22　Serial突入Highリミット
    Dim dblHiLoLimValid As Long             'Arg23　Serial突入リミットの有効範囲
    
    '変数取り込み
    If Not Sub_SeriParaJudge_GetParameter( _
                strResultKey, _
                dblLoJudgeLimit, _
                dblHiJudgeLimit, _
                dblHiLoLimValid) Then
                MsgBox "The Number of subCurrentSeriParaJudge_f's arguments is invalid!"
                Call DisableAllTest 'EeeJob関数
                Exit Function
    End If
    
    'SeriParaJudge
    Call mf_Sub_SeriParaJudge(strResultKey, dblLoJudgeLimit, dblHiJudgeLimit, dblHiLoLimValid, Go_Serial_Mesure)
    
End Function


'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Parallel→Serial_Test用:Start(パラレル測定側)
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
'2013/12/05 T.Morimoto
'               神山さんが、IMX219において追加項目としてHWスタンバイでのシリアル測定
'               (パラレル測定後判定→シリアル測定)が導入されたことに伴い追加。
'               　BPMUとそれ以外での分岐ができていないため、その分岐を追記。
Private Function SubCurrentTestIfNeeded_f() As Double

    On Error GoTo ErrorExit
    
    '結果変数
    Dim retResult() As Double                   '2013/02/05 修正
    Dim retResult2(nSite) As Double             '2013/02/05 修正
    '本来のテストラベルを作る。測定要求仕様書に記載されているテストラベル名は、
    '本テストインスタンスの"__"ならびにこれに続く文字列を除外することで得られる。
    '   To obtain the original test label on the specification sheet. It can be
    'obtained by removing string starting with "__" from this instance name.
    Dim strResultKey As String
    strResultKey = UCase(TheExec.DataManager.InstanceName)
    
    If Go_Serial_Mesure = True Then

        'これひとつでひとつのテストするのでSiteCheckは必要
        Call SiteCheck
        
        '変数定義
        Dim strPin As String                    'Arg20　テスト端子 (Test pin name)
        Dim dblForceVoltage As Double           'Arg21　印加電圧 (VDD bias value)
        Dim strSetParamCondition As String      'Arg22-1　測定パラメータ_Opt_リレー (パラレル測定結果)
        Dim strPowerCondition As String         'Arg22-2　PPS & Pin settings
        Dim strPatternCondition As String       'Arg22-3　Pattern
            
        '測定パラメータ
        Dim lAve As Double                      'TestCondition
        Dim dblClampCurrent As Double           'TestCondition
        Dim dblWait As Double                   'TestCondition
        Dim dblPinResourceName As String        'TestCondition
                
        '関数内変数
        Dim Flg_Active(nSite) As Long
        Dim TempValue(nSite) As Double
        Dim site As Long
        Dim mychanType As chtype
        
        '変数取り込み
        '   To obtain the argument parameters on test instances sheet.
        If Not getParam_SerialMeasureAfterParallel( _
                  strPin, _
                  dblForceVoltage, _
                  strSetParamCondition, _
                  strPowerCondition, _
                  strPatternCondition) Then
                  MsgBox "The Number of subCurrent_Test_f's arguments is invalid!"
                  Call DisableAllTest 'EeeJob関数
                  Exit Function
        End If
            
        'パラメータ設定の関数を呼ぶ (FW_SetSubCurrentParam)
        '   To call measurement parameter setting condition and environment setup condition.
        Call TheCondition.SetCondition(strSetParamCondition) '1
    
        'Test Conditionで設定されているパラメータをVarBankより取得
        '   To obtain dc measurement parameters registered in the VarBank.
        lAve = GetSubCurrentAverageCount(GetInstanceName)
        dblClampCurrent = GetSubCurrentClampCurrent(GetInstanceName)
        dblWait = GetSubCurrentWaitTime(GetInstanceName)
        dblPinResourceName = GetSubCurrentPinResourceName(GetInstanceName)
    
        'Activeサイトの確認
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                Flg_Active(site) = 1
            End If
        Next site
        
        'チャネルタイプの確認
        mychanType = GetChanType(strPin)

        'SUB電流測定の確認
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                '========== 未使用siteのBetaGND切り離し ===============================
                Call GndSeparateBySite(site)
                
                '========== Set Condition ===============================
                '   To execute device setup (power and pin electronics setting)
                Call TheCondition.SetCondition(strPowerCondition) '2
    
                '========== Force Voltage ===============================
                '測定。HSD200であってもBPMUを使います。
                'Measurement. Measurement will be performed with BPMU in case of digital channel,
                'regardless of HSD100 or HSD200.
                If mychanType = chIO Then
                    Call SetFVMI_BPMU(strPin, 0 * V, dblClampCurrent, site)
                    Call DisconnectPins(strPin, site)
                    Call SetFVMI_BPMU(strPin, dblForceVoltage, dblClampCurrent, site)
                Else
                    '========== Force Voltage ===============================
                    Call SetFVMI(strPin, 0 * V, dblClampCurrent, site)
                    Call DisconnectPins(strPin, site)
                    Call SetFVMI(strPin, dblForceVoltage, dblClampCurrent, site)
                End If
                
                '========== Set Pattern ===============================
                Call TheCondition.SetCondition(strPatternCondition) '3
                TheHdw.WAIT dblWait
                    
                '========== Measure Current ===============================
                If mychanType = chIO Then
                    Call MeasureI_BPMU(strPin, retResult2, lAve, site)
                    Call SetFVMI_BPMU(strPin, 0 * V, dblClampCurrent, site)
                    Call DisconnectPins(strPin, site)
                Else
                    Call MeasureI(strPin, retResult2, lAve, site)
                    Call SetFVMI(strPin, 0 * V, dblClampCurrent, site)
                    Call DisconnectPins(strPin, site)
                End If
                
                '========== 未使用siteのBetaGND戻し =======================
                Call GndConectBySite(site, Flg_Active)
            End If
        Next site
        
      
        'パターン停止
        Call StopPattern 'EeeJob関数
      
        'All_Open及びDisconnect自動生成化
        Call PowerDownAndDisconnect
                
        '答えは返さずAddするのみ; Add the result to the Result Manager.
        Call updateResult(strResultKey, retResult2)
        Call test(retResult2)
    Else
        Call TheResult.GetResult(strResultKey, retResult)
        Call test(retResult)
    End If
    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call TheResult.Add(strResultKey, retResult)

End Function


'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Parallel→Serial_Test用:Start
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
' History:  First drafted by T.Koyama 2013-12-04 (IMX219 MP1)
'           Site serial current measurement method without pattern burst.
Private Function SubCurrentTestIfNeededNoPattern_f() As Double

    On Error GoTo ErrorExit
    
    '結果変数
    Dim retResult() As Double                   '2013/02/05 修正
    Dim retResult2(nSite) As Double             '2013/02/05 修正
    '本来のテストラベルを作る。測定要求仕様書に記載されているテストラベル名は、
    '本テストインスタンスの"__"ならびにこれに続く文字列を除外することで得られる。
    '   To obtain the original test label on the specification sheet. It can be
    'obtained by removing string starting with "__" from this instance name.
    Dim strResultKey As String
    strResultKey = UCase(TheExec.DataManager.InstanceName)
    
    If Go_Serial_Mesure = True Then

        'これひとつでひとつのテストするのでSiteCheckは必要
        Call SiteCheck
        
        '変数定義
        Dim strPin As String                    'Arg20　テスト端子 (Test pin name)
        Dim dblForceVoltage As Double           'Arg21　印加電圧 (VDD bias value)
        Dim strSetParamCondition As String      'Arg22-1　測定パラメータ_Opt_リレー (パラレル測定結果)
        Dim strPowerCondition As String         'Arg22-2　PPS & Pin settings
            
        '測定パラメータ
        Dim lAve As Double                      'TestCondition
        Dim dblClampCurrent As Double           'TestCondition
        Dim dblWait As Double                   'TestCondition
        Dim dblPinResourceName As String        'TestCondition
                
        '関数内変数
        Dim Flg_Active(nSite) As Long
        Dim TempValue(nSite) As Double
        Dim site As Long
        Dim mychanType As chtype
        
        '変数取り込み
        '   To obtain the argument parameters on test instances sheet.
        If Not getParam_SerialMeasureAfterParallel_NoPattern( _
                  strPin, _
                  dblForceVoltage, _
                  strSetParamCondition, _
                  strPowerCondition) Then
                  MsgBox "The Number of subCurrent_Test_f's arguments is invalid!"
                  Call DisableAllTest 'EeeJob関数
                  Exit Function
        End If
            
        'パラメータ設定の関数を呼ぶ (FW_SetSubCurrentParam)
        '   To call measurement parameter setting condition and environment setup condition.
        Call TheCondition.SetCondition(strSetParamCondition) '1
    
        'Test Conditionで設定されているパラメータをVarBankより取得
        '   To obtain dc measurement parameters registered in the VarBank.
        lAve = GetSubCurrentAverageCount(GetInstanceName)
        dblClampCurrent = GetSubCurrentClampCurrent(GetInstanceName)
        dblWait = GetSubCurrentWaitTime(GetInstanceName)
        dblPinResourceName = GetSubCurrentPinResourceName(GetInstanceName)
    
        'チャネルタイプの確認
        mychanType = GetChanType(strPin)
        
        'Activeサイトの確認
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                Flg_Active(site) = 1
            End If
        Next site
    
        'SUB電流測定の確認
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                '========== 未使用siteのBetaGND切り離し ===============================
                Call GndSeparateBySite(site)
                
                '========== Set Condition ===============================
                '   To execute device setup (power and pin electronics setting)
                Call TheCondition.SetCondition(strPowerCondition) '2

                '========== Force Voltage ===============================
                '測定。HSD200であってもBPMUを使います。
                'Measurement. Measurement will be performed with BPMU in case of digital channel,
                'regardless of HSD100 or HSD200.
                'Digital ChannelならBPMUを使用する。但し、HSD200であってもBPMUを使用する(IMX164の相関から必要と判断された)
                If mychanType = chIO Then
                    Call SetFVMI_BPMU(strPin, 0 * V, dblClampCurrent, site)
                    Call DisconnectPins(strPin, site)
                    Call SetFVMI_BPMU(strPin, dblForceVoltage, dblClampCurrent, site)
                Else
                    Call SetFVMI(strPin, 0 * V, dblClampCurrent, site)
                    Call DisconnectPins(strPin, site)
                    Call SetFVMI(strPin, dblForceVoltage, dblClampCurrent, site)
                End If
                
                '========== Set Pattern ===============================
                TheHdw.WAIT dblWait
                
                '========== Measure Current ===============================
                If mychanType = chIO Then
                    Call MeasureI_BPMU(strPin, retResult2, lAve, site)
                    Call SetFVMI_BPMU(strPin, 0 * V, dblClampCurrent, site)
                    Call DisconnectPins(strPin, site)
                Else
                    Call MeasureI(strPin, retResult2, lAve, site)
                    Call SetFVMI(strPin, 0 * V, dblClampCurrent, site)
                    Call DisconnectPins(strPin, site)
                End If
    
                '========== 未使用siteのBetaGND戻し =======================
                Call GndConectBySite(site, Flg_Active)
            End If
        Next site
     
        'All_Open及びDisconnect自動生成化
        Call PowerDownAndDisconnect
                
        '答えは返さずAddするのみ; Add the result to the Result Manager.
        Call updateResult(strResultKey, retResult2)
        Call test(retResult2)
    Else
        Call TheResult.GetResult(strResultKey, retResult)
        Call test(retResult)
    End If
    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call TheResult.Add(strResultKey, retResult)

End Function

Private Function updateResult(ByVal keyName As String, ByRef resultValue() As Double) As Boolean
    On Error GoTo ErrorDetected
    Call TheResult.Add(keyName, resultValue)
    updateResult = True
    Exit Function
ErrorDetected:
    Call TheResult.Delete(keyName)
    Call TheResult.Add(keyName, resultValue)
    updateResult = True
End Function

Private Function Sub_SeriParaJudge_GetParameter( _
    ByRef strResultKey As String, _
    ByRef dblLoJudgeLimit As Double, _
    ByRef dblHiJudgeLimit As Double, _
    ByRef dblHiLoLimValid As Long _
    ) As Boolean

On Error GoTo ErrHandler
    '変数取り込み
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_SUB_SERIPARA_JUDGE_ARGS) Then
        Sub_SeriParaJudge_GetParameter = False
        Exit Function
    End If
    
    strResultKey = ArgArr(0)
    
    If ArgArr(1) <> "" Then
        dblLoJudgeLimit = ArgArr(1)
    Else
        ArgArr(1) = 0
    End If
    
    If ArgArr(2) <> "" Then
        dblHiJudgeLimit = ArgArr(2)
    Else
        ArgArr(2) = 0
    End If
    
    dblHiLoLimValid = ArgArr(3)
    On Error GoTo 0

    Sub_SeriParaJudge_GetParameter = True
    Exit Function
    
ErrHandler:

    Sub_SeriParaJudge_GetParameter = False
    Exit Function

End Function

Public Sub mf_Sub_SeriParaJudge(ByVal Data As String, ByRef LoLimit As Double, ByRef HiLimit As Double, ByRef HiLoLimValid As Long, ByRef Retry_flag As Boolean) '2012/11/16 175Debug Arikawa
'方法:
'   1. パラレル測定をする
'   2. パラレル測定の結果を用いて平均値を求める。
'
'シリアル→パラレル判定
'   シリパラ判定の突入条件値を以下の方法で計算される。
'   1. N増し測定した結果のヒストグラムにおいて、分布の最も厳しい側(低い側)と判断
'      される値を決める<分布下限値>。(尚、ここでいう分布とは、「方法」の2に記載した
'       「平均値」の分布のこと。
'   2. 指定される<スペック値>を用いて、突入条件値は以下の式で得られる。
'           <突入条件値> = <分布下限値> * <サイト数> + (<スペック値> - <分布下限値>)
'                          ~~~~~~~~~~~~~~~~~~~~~~~~    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'       右辺第一項だけだと、分布下限値なので、ほとんど常にシリアル測定に突入していしまう。
'       そこでマージンとして用意されているのが右辺第二項。
'   3. このようにして決められた突入条件値は、パラレル測定結果の平均値をサイト数倍した
'       値に適用しなければならない。パラレル測定結果の合計値(生きているサイト数倍したもの)
'      に適用するのはだめ。
'注意
'   理収外コンタクト時には必ずシリアル測定を実施する
'
'Method:
'   1. Parallel current measurement.
'   2. Calculating the mean value of step1, judge if serial measurement is needed.
'
'Judgement
'   The limit value for the judgement is calculated based on the following idea.
'   1. Using a large amount of chip measurement result, determine a value at the lower bound
'      of measurement value histogram. (Each measurement result is the mean
'      value described at step 2 of "Method" section.
'   2. Assume that the spec value <spec>, the limit value is calculated by the equation
'             <limit value> = <lower bound value> * <nSite + 1> + (<spec> - <lower bound value>)
'                             ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'        The first term on the right only will result in execution of serial measurement
'       in almost all the cases, since <lower bound value> is at the lower bound of histogram.
'       The second term on the right, therefore, is the margin term.
'   3. The <limit value> must be applied to the measured mean value multiplied by the
'      number of sites under test (not the number of sites alive in judgement).
'Notice:
'   At least more than one site is located beyond wafer edge, serial current measurement
'   must be unconditionally executed.

    Dim site As Long
    Dim Active_site As Long
    Dim TempValue() As Double
    Dim tempValueAve(nSite) As Double
    Dim tempValueSum As Double
                
    'パラレル測定結果を取得(DC Test ScenarioのOperation-Resultに格納)
    'To obtain parallel measurement result (conducted by DC Test Scenario)(stored in Operation-Result).
    Call TheDcTest.GetTempResult(Data, TempValue)
'    Call TheResult.GetResult(Data, tempValue) '2012/11/16 175Debug Arikawa
    
    'テストラベルを分解し、本来格納すべきテストラベルを抽出する。入力されるテストラベルは、
    '測定要求仕様書に指定されるテストラベルに対して、"__"とこれに続く文字列が追加されている。
    'To obtain the test label by extracting the input operation-result label name, assuming
    'that the operation-result label is made with a string starting with "__" following
    'the test label on the specification sheet.
    Dim uScorePos As Long
    Dim originalTestLabel As String
    uScorePos = InStr(Data, "__")
    originalTestLabel = Mid(Data, 1, uScorePos - 1)
    
    '========== RETRY CHECK ======================================
    Retry_flag = False

    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tempValueSum = tempValueSum + TempValue(site)
            Active_site = Active_site + 1
        End If
    Next site

    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tempValueAve(site) = tempValueSum / Active_site
        End If
    Next site

'    Call TheResult.Delete(Data) '2012/11/16 175Debug Arikawa
    '答えは返さずAddするのみ
'    Call TheResult.Add(Data, tempValueAve)  '2012/11/16 175Debug Arikawa
    Call TheResult.Add(originalTestLabel, tempValueAve)
    
'    'シリアル測定突入条件値の算出
'    tempValueSum = tempValueSum / Active_site * SITE_MAX
    
    Select Case HiLoLimValid
        Case 0
            If Active_site < SITE_MAX Then
                Retry_flag = True
            End If
        Case 1
            If tempValueSum < LoLimit Or Active_site < SITE_MAX Then
                Retry_flag = True
            End If
        Case 2
            If tempValueSum > HiLimit Or Active_site < SITE_MAX Then
                Retry_flag = True
            End If
        Case 3
            If (tempValueSum < LoLimit Or tempValueSum > HiLimit) Or Active_site < SITE_MAX Then
                Retry_flag = True
            End If
        Case Else
    End Select
    
End Sub

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Parallel→Serial_Test用:End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Serial_NoPattern_Test用:Start
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'2013/02/07
'♪
Private Function subCurrent_Serial_NoPattern_Test_f() As Double

    On Error GoTo ErrorExit

    'これひとつでひとつのテストするのでSiteCheckは必要
    Call SiteCheck
        
    '変数定義
    Dim strResultKey As String              'Arg20　項目名
    Dim strPin As String                    'Arg21　テスト端子
    Dim dblForceVoltage As Double           'Arg22　印加電圧
    Dim strSetParamCondition As String      'Arg23　測定パラメータ_Opt_リレー
    Dim strPowerCondition As String         'Arg24　Set_Voltage_端子設定
'    Dim strPatternCondition As String       'Arg25　Pattern
            
    '測定パラメータ
    Dim lAve As Double                      'TestCondition
    Dim dblClampCurrent As Double           'TestCondition
    Dim dblWait As Double                   'TestCondition
    Dim dblPinResourceName As String        'TestCondition
        
    '結果変数
    Dim retResult(nSite) As Double
            
    '関数内変数
    Dim Flg_Active(nSite) As Long
    Dim TempValue(nSite) As Double
    Dim site As Long
            
    '変数取り込み
    If Not SubCurrentTest_NoPattern_GetParameter( _
                strResultKey, _
                strPin, _
                dblForceVoltage, _
                strSetParamCondition, _
                strPowerCondition) Then
                MsgBox "The Number of subCurrent_Serial_NoPattern_Test_f's arguments is invalid!"
                Call DisableAllTest 'EeeJob関数
                Exit Function
    End If
            
    'パラメータ設定の関数を呼ぶ (FW_SetSubCurrentParam)
    Call TheCondition.SetCondition(strSetParamCondition)

    lAve = GetSubCurrentAverageCount(GetInstanceName)
    dblClampCurrent = GetSubCurrentClampCurrent(GetInstanceName)
    dblWait = GetSubCurrentWaitTime(GetInstanceName)
    dblPinResourceName = GetSubCurrentPinResourceName(GetInstanceName)

    'Activeサイトの確認
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Flg_Active(site) = 1
        End If
    Next site

    'SUB電流測定の確認
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            '========== 未使用siteのBetaGND切り離し ===============================
            Call GndSeparateBySite(site)
            
            '========== Set Condition ===============================
            Call TheCondition.SetCondition(strPowerCondition)

             '========== Force Voltage ===============================
            If StrComp(dblPinResourceName, "BPMU", vbTextCompare) = 0 Then
                Call SetFVMI_BPMU(strPin, 0 * V, dblClampCurrent, site)
                Call DisconnectPins(strPin, site)
                Call SetFVMI_BPMU(strPin, dblForceVoltage, dblClampCurrent, site)
            Else
                Call SetFVMI(strPin, 0 * V, dblClampCurrent, site)
                Call DisconnectPins(strPin, site)
                Call SetFVMI(strPin, dblForceVoltage, dblClampCurrent, site)
            End If
            
            TheHdw.WAIT dblWait
            
             '========== Measure Current ===============================
            If StrComp(dblPinResourceName, "BPMU", vbTextCompare) = 0 Then
                Call MeasureI_BPMU(strPin, retResult, lAve, site)
                Call SetFVMI_BPMU(strPin, 0 * V, dblClampCurrent, site)
                Call DisconnectPins(strPin, site)
            Else
                Call MeasureI(strPin, retResult, lAve, site)
                Call SetFVMI(strPin, 0 * V, dblClampCurrent, site)
                Call DisconnectPins(strPin, site)
            End If

             '========== 未使用siteのBetaGND戻し =======================
            Call GndConectBySite(site, Flg_Active)
        End If
    Next site
    
    'パターン停止
    Call StopPattern 'EeeJob関数
  
    'All_Open及びDisconnect自動生成化
    Call PowerDownAndDisconnect
           
    '答えも返しAddもする
    Call test(retResult)
    Call TheResult.Add(strResultKey, retResult)

    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call test(retResult)
    Call TheResult.Add(strResultKey, retResult)

End Function
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'SubCurrentTest_NoPattern_GetParameter用
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
'2013/02/06
Private Function SubCurrentTest_NoPattern_GetParameter( _
    ByRef strResultKey As String, _
    ByRef strPin As String, _
    ByRef dblForceVoltage As Double, _
    ByRef strSetParamCondition As String, _
    ByRef strPowerCondition As String _
    ) As Boolean

    '変数取り込み
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_SUBCURRENT_ARGS) Then
        SubCurrentTest_NoPattern_GetParameter = False
        Exit Function
    End If
    
    Dim tempstr As String      '一時保存変数
    Dim tempArrstr() As String '一時保存配列
    
On Error GoTo ErrHandler
    strResultKey = ArgArr(0)                'Arg20: Test label name.
    strPin = ArgArr(1)                      'Arg21: Test pin name
    dblForceVoltage = CDbl(ArgArr(2))       'Arg22: Force voltage (PPS value) for test pin.
    strSetParamCondition = ArgArr(3)        'Arg23: [Test Condition]'s condition name for common environment setup.
    strPowerCondition = ArgArr(4)           'Arg24: [Test Condition]'s condition name for device setup.
On Error GoTo 0

    SubCurrentTest_NoPattern_GetParameter = True
    Exit Function
    
ErrHandler:

    SubCurrentTest_NoPattern_GetParameter = False
    Exit Function

End Function


'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Parallel→BPMU Serial_Judge用:Start
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

Private Function subCurrentNonScenarioSeriParaJudge_f() As Double

    'これひとつでひとつのテストするのでSiteCheckは必要
    Call SiteCheck
    
    Go_Serial_Mesure = False
    
    '変数定義
    Dim strResultKey As String              'Arg20　Test label
    Dim dblLoJudgeLimit As Double           'Arg21　Serial突入Lowリミット
    Dim dblHiJudgeLimit As Double           'Arg22　Serial突入Highリミット
    Dim dblHiLoLimValid As Long             'Arg23　Serial突入リミットの有効範囲
    

    Dim temp_strResultKey As String
    temp_strResultKey = UCase(TheExec.DataManager.InstanceName)
    Dim uScorePos As Long
    Dim originalTestLabel As String
    uScorePos = InStr(temp_strResultKey, "__")
    originalTestLabel = Mid(temp_strResultKey, 1, uScorePos - 1)
    strResultKey = originalTestLabel & "__SeriParaParaMeasure"
    
    
    '変数取り込み
    If Not Sub_NonScenarioParaJudge_GetParameter( _
                dblLoJudgeLimit, _
                dblHiJudgeLimit, _
                dblHiLoLimValid) Then
                MsgBox "The Number of subCurrentSeriParaJudge_f's arguments is invalid!"
                Call DisableAllTest 'EeeJob関数
                Exit Function
    End If
    
    'SeriParaJudge
    Call mf_Sub_SeriParaBPMUJudge(strResultKey, dblLoJudgeLimit, dblHiJudgeLimit, dblHiLoLimValid, Go_Serial_Mesure)
    
End Function

Public Sub mf_Sub_SeriParaBPMUJudge(ByVal Data As String, ByRef LoLimit As Double, ByRef HiLimit As Double, ByRef HiLoLimValid As Long, ByRef Retry_flag As Boolean) '2012/11/16 175Debug Arikawa
'方法:
'   1. パラレル測定をする
'   2. パラレル測定の結果を用いて平均値を求める。
'
'シリアル→パラレル判定
'   シリパラ判定の突入条件値を以下の方法で計算される。
'   1. N増し測定した結果のヒストグラムにおいて、分布の最も厳しい側(低い側)と判断
'      される値を決める<分布下限値>。(尚、ここでいう分布とは、「方法」の2に記載した
'       「平均値」の分布のこと。
'   2. 指定される<スペック値>を用いて、突入条件値は以下の式で得られる。
'           <突入条件値> = <分布下限値> * <サイト数> + (<スペック値> - <分布下限値>)
'                          ~~~~~~~~~~~~~~~~~~~~~~~~    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'       右辺第一項だけだと、分布下限値なので、ほとんど常にシリアル測定に突入していしまう。
'       そこでマージンとして用意されているのが右辺第二項。
'   3. このようにして決められた突入条件値は、パラレル測定結果の平均値をサイト数倍した
'       値に適用しなければならない。パラレル測定結果の合計値(生きているサイト数倍したもの)
'      に適用するのはだめ。
'注意
'   理収外コンタクト時には必ずシリアル測定を実施する
'
'Method:
'   1. Parallel current measurement.
'   2. Calculating the mean value of step1, judge if serial measurement is needed.
'
'Judgement
'   The limit value for the judgement is calculated based on the following idea.
'   1. Using a large amount of chip measurement result, determine a value at the lower bound
'      of measurement value histogram. (Each measurement result is the mean
'      value described at step 2 of "Method" section.
'   2. Assume that the spec value <spec>, the limit value is calculated by the equation
'             <limit value> = <lower bound value> * <nSite + 1> + (<spec> - <lower bound value>)
'                             ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'        The first term on the right only will result in execution of serial measurement
'       in almost all the cases, since <lower bound value> is at the lower bound of histogram.
'       The second term on the right, therefore, is the margin term.
'   3. The <limit value> must be applied to the measured mean value multiplied by the
'      number of sites under test (not the number of sites alive in judgement).
'Notice:
'   At least more than one site is located beyond wafer edge, serial current measurement
'   must be unconditionally executed.

    Dim site As Long
    Dim Active_site As Long
    Dim TempValue() As Double
    Dim tempValueAve(nSite) As Double
    Dim tempValueSum As Double
                
    'パラレル測定結果を取得(DC Test ScenarioのOperation-Resultに格納)
    'To obtain parallel measurement result (conducted by DC Test Scenario)(stored in Operation-Result).
'    Call TheDcTest.GetTempResult(Data, tempValue)
    Call TheResult.GetResult(Data, TempValue) '2012/11/16 175Debug Arikawa
    
    'テストラベルを分解し、本来格納すべきテストラベルを抽出する。入力されるテストラベルは、
    '測定要求仕様書に指定されるテストラベルに対して、"__"とこれに続く文字列が追加されている。
    'To obtain the test label by extracting the input operation-result label name, assuming
    'that the operation-result label is made with a string starting with "__" following
    'the test label on the specification sheet.
    Dim uScorePos As Long
    Dim originalTestLabel As String
    uScorePos = InStr(Data, "__")
    originalTestLabel = Mid(Data, 1, uScorePos - 1)
    
    '========== RETRY CHECK ======================================
    Retry_flag = False
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tempValueSum = tempValueSum + TempValue(site)
            Active_site = Active_site + 1
        End If
    Next site

    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tempValueAve(site) = tempValueSum / Active_site
        End If
    Next site

'    Call TheResult.Delete(Data) '2012/11/16 175Debug Arikawa
    '答えは返さずAddするのみ
'    Call TheResult.Add(Data, tempValueAve)  '2012/11/16 175Debug Arikawa
    Call TheResult.Add(originalTestLabel, tempValueAve)
    
'    'シリアル測定突入条件値の算出
'    tempValueSum = tempValueSum / Active_site * SITE_MAX
    
    Select Case HiLoLimValid
        Case 0
            If Active_site < SITE_MAX Then
                Retry_flag = True
            End If
        Case 1
            If tempValueSum < LoLimit Or Active_site < SITE_MAX Then
                Retry_flag = True
            End If
        Case 2
            If tempValueSum > HiLimit Or Active_site < SITE_MAX Then
                Retry_flag = True
            End If
        Case 3
            If (tempValueSum < LoLimit Or tempValueSum > HiLimit) Or Active_site < SITE_MAX Then
                Retry_flag = True
            End If
        Case Else
    End Select
    
End Sub

'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Paralle  BPMU 2013/3/7  Hamada
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

Private Function SubCurrentNonScenario_Measure_f() As Double

    On Error GoTo ErrorExit
    
    '結果変数
    Dim retResult() As Double                   '2013/02/05 修正
    Dim retResult2(nSite) As Double             '2013/02/05 修正
    '本来のテストラベルを作る。測定要求仕様書に記載されているテストラベル名は、
    '本テストインスタンスの"__"ならびにこれに続く文字列を除外することで得られる。
    '   To obtain the original test label on the specification sheet. It can be
    'obtained by removing string starting with "__" from this instance name.
    Dim temp_strResultKey As String
    Dim strResultKey As String
    temp_strResultKey = UCase(TheExec.DataManager.InstanceName)
'    strResultKey = UCase(TheExec.DataManager.instanceName)
    Dim uScorePos As Long
    Dim originalTestLabel As String
    uScorePos = InStr(temp_strResultKey, "__")
    originalTestLabel = Mid(temp_strResultKey, 1, uScorePos - 1)
    strResultKey = originalTestLabel & "__SeriParaParaMeasure"

        'これひとつでひとつのテストするのでSiteCheckは必要
    Call SiteCheck
    
    '変数定義
    Dim strPin As String                    'Arg20　テスト端子 (Test pin name)
    Dim dblForceVoltage As Double           'Arg21　印加電圧 (VDD bias value)
    Dim strSetParamCondition As String      'Arg22-1　測定パラメータ_Opt_リレー (パラレル測定結果)
    Dim strPowerCondition As String         'Arg22-2　PPS & Pin settings
    Dim strPatternCondition As String       'Arg22-3　Pattern
        
    '測定パラメータ
    Dim lAve As Double                      'TestCondition
    Dim dblClampCurrent As Double           'TestCondition
    Dim dblWait As Double                   'TestCondition
    Dim dblPinResourceName As String        'TestCondition
            
    '関数内変数
    Dim Flg_Active(nSite) As Long
'    Dim tempValue(nSite) As Double
'    Dim strResultKey As String
    Dim site As Long
        
    Dim mychanType As chtype
    Dim FunctionName As String
    '変数取り込み
    '   To obtain the argument parameters on test instances sheet.
    If Not getParam_BPMU_Parallel( _
              strPin, _
              dblForceVoltage, _
              strSetParamCondition, _
              strPowerCondition, _
              strPatternCondition _
              ) Then
              MsgBox "The Number of subCurrent_Test_f's arguments is invalid!"
              Call DisableAllTest 'EeeJob関数
              Exit Function
    End If
    
    Dim Temp_ChanArr() As Long
    Dim Temp_chanCnt As Long
    Dim Temp_BoardCnt As Long
    Dim Temp_ChanbyBoard As Long
    Dim Temp_Message As String

    Dim tempValueSum As Double
    Dim TempValue(nSite) As Double
    
    Dim i As Long
    Dim ii As Long
    Dim myboard_i As Long
    myboard_i = 0
    
    
    Dim Active_site As Long
    Active_site = 0
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            Active_site = Active_site + 1
        End If
    Next site
    
    
    If TesterType <> "IP750EX" Then
        
        mychanType = GetChanType(strPin)
        Call TheExec.DataManager.GetChanListByBoard(strPin, ALL_SITE, mychanType, Temp_ChanArr, Temp_chanCnt, Temp_BoardCnt, Temp_ChanbyBoard, Temp_Message)
        
        Call TheCondition.SetCondition(strSetParamCondition) '1
    
        'Test Conditionで設定されているパラメータをVarBankより取得
        lAve = GetSubCurrentAverageCount(GetInstanceName)
        dblClampCurrent = GetSubCurrentClampCurrent(GetInstanceName)
        dblWait = GetSubCurrentWaitTime(GetInstanceName)
        dblPinResourceName = GetSubCurrentPinResourceName(GetInstanceName)
        
             If mychanType = chIO Then
'                Dim cBoard As Long
'                For cBoard = 0 To Temp_BoardCnt - 1

                '========== Set Condition ===============================
                Call TheCondition.SetCondition(strPowerCondition) '2
        
                '========== Force Voltage ===============================
                Call SetFVMI_BPMU(strPin, dblForceVoltage, dblClampCurrent)
        
                '========== Set Pattern ===============================
                Call TheCondition.SetCondition(strPatternCondition) '3
                TheHdw.WAIT dblWait
                
                '========== Measure Current ===============================
                Call MeasureI_BPMU(strPin, TempValue, lAve)
                                
                For i = 0 To Temp_BoardCnt - 1
                        tempValueSum = tempValueSum + TempValue(myboard_i)
                        myboard_i = myboard_i + Temp_ChanbyBoard
                Next i
                
                For site = 0 To nSite
                    If TheExec.sites.site(site).Active = True Then
                            retResult2(site) = tempValueSum
                    End If
                Next site
                
            Else
                '========== Set Condition ===============================
                Call TheCondition.SetCondition(strPowerCondition) '2
                '========== Force Voltage ===============================
                Call SetFVMI(strPin, dblForceVoltage, dblClampCurrent)
                '========== Set Pattern ===============================
                Call TheCondition.SetCondition(strPatternCondition) '3
                TheHdw.WAIT dblWait
                '========== Measure Current ===============================
                Call MeasureI(strPin, retResult2, lAve)
            End If
         
            
        'パターン停止
        Call StopPattern 'EeeJob関数
        
        Call DisconnectPins(strPin, ALL_SITE)
        'All_Open及びDisconnect自動生成化
        Call PowerDownAndDisconnect
        
    Else
    
    
        Call TheCondition.SetCondition(strSetParamCondition) '1
        lAve = GetSubCurrentAverageCount(GetInstanceName)
        dblClampCurrent = GetSubCurrentClampCurrent(GetInstanceName)
        dblWait = GetSubCurrentWaitTime(GetInstanceName)
        dblPinResourceName = GetSubCurrentPinResourceName(GetInstanceName)
        
        '========== Set Condition ===============================
        Call TheCondition.SetCondition(strPowerCondition) '2
        '========== Force Voltage ===============================
        Call SetFVMI(strPin, dblForceVoltage, dblClampCurrent)
        '========== Set Pattern ===============================
        Call TheCondition.SetCondition(strPatternCondition) '3
        TheHdw.WAIT dblWait
        '========== Measure Current ===============================
        Call MeasureI(strPin, retResult2, lAve, site)
            
        'パターン停止
        Call StopPattern 'EeeJob関数
        Call DisconnectPins(strPin, ALL_SITE)
        Call PowerDownAndDisconnect
        
    End If

    Call updateResult(strResultKey, retResult2)

    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call TheResult.Add(strResultKey, retResult)

End Function
'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Paralle  BPMU 2013/3/7  Hamada
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪

Private Function SubCurrentNonScenarioNoPattern_Measure_f() As Double

    On Error GoTo ErrorExit
    
    '結果変数
    Dim retResult() As Double                   '2013/02/05 修正
    Dim retResult2(nSite) As Double             '2013/02/05 修正
    '本来のテストラベルを作る。測定要求仕様書に記載されているテストラベル名は、
    '本テストインスタンスの"__"ならびにこれに続く文字列を除外することで得られる。
    '   To obtain the original test label on the specification sheet. It can be
    'obtained by removing string starting with "__" from this instance name.
    Dim temp_strResultKey As String
    Dim strResultKey As String
    temp_strResultKey = UCase(TheExec.DataManager.InstanceName)
'    strResultKey = UCase(TheExec.DataManager.instanceName)
    Dim uScorePos As Long
    Dim originalTestLabel As String
    uScorePos = InStr(temp_strResultKey, "__")
    originalTestLabel = Mid(temp_strResultKey, 1, uScorePos - 1)
    strResultKey = originalTestLabel & "__SeriParaParaMeasure"

        'これひとつでひとつのテストするのでSiteCheckは必要
    Call SiteCheck
    
    '変数定義
    Dim strPin As String                    'Arg20　テスト端子 (Test pin name)
    Dim dblForceVoltage As Double           'Arg21　印加電圧 (VDD bias value)
    Dim strSetParamCondition As String      'Arg22-1　測定パラメータ_Opt_リレー (パラレル測定結果)
    Dim strPowerCondition As String         'Arg22-2　PPS & Pin settings
    Dim strPatternCondition As String       'Arg22-3　Pattern
        
    '測定パラメータ
    Dim lAve As Double                      'TestCondition
    Dim dblClampCurrent As Double           'TestCondition
    Dim dblWait As Double                   'TestCondition
    Dim dblPinResourceName As String        'TestCondition
            
    '関数内変数
    Dim Flg_Active(nSite) As Long
'    Dim tempValue(nSite) As Double
'    Dim strResultKey As String
    Dim site As Long
        
    Dim mychanType As chtype
    Dim FunctionName As String
    '変数取り込み
    '   To obtain the argument parameters on test instances sheet.
    If Not getParam_BPMU_Parallel_NonPattern( _
              strPin, _
              dblForceVoltage, _
              strSetParamCondition, _
              strPowerCondition _
              ) Then
              MsgBox "The Number of subCurrent_Test_f's arguments is invalid!"
              Call DisableAllTest 'EeeJob関数
              Exit Function
    End If
    
    Dim Temp_ChanArr() As Long
    Dim Temp_chanCnt As Long
    Dim Temp_BoardCnt As Long
    Dim Temp_ChanbyBoard As Long
    Dim Temp_Message As String

    Dim tempValueSum As Double
    Dim TempValue(nSite) As Double
    
    Dim i As Long
    Dim ii As Long
    Dim myboard_i As Long
    myboard_i = 0
    
    
    Dim Active_site As Long
    Active_site = 0
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            Active_site = Active_site + 1
        End If
    Next site
    
    
    If TesterType <> "IP750EX" Then
        
        mychanType = GetChanType(strPin)
        Call TheExec.DataManager.GetChanListByBoard(strPin, ALL_SITE, mychanType, Temp_ChanArr, Temp_chanCnt, Temp_BoardCnt, Temp_ChanbyBoard, Temp_Message)
        
        Call TheCondition.SetCondition(strSetParamCondition) '1
    
        'Test Conditionで設定されているパラメータをVarBankより取得
        lAve = GetSubCurrentAverageCount(GetInstanceName)
        dblClampCurrent = GetSubCurrentClampCurrent(GetInstanceName)
        dblWait = GetSubCurrentWaitTime(GetInstanceName)
        dblPinResourceName = GetSubCurrentPinResourceName(GetInstanceName)
        
             If mychanType = chIO Then
'                Dim cBoard As Long
'                For cBoard = 0 To Temp_BoardCnt - 1

                '========== Set Condition ===============================
                Call TheCondition.SetCondition(strPowerCondition) '2
        
                '========== Force Voltage ===============================
                Call SetFVMI_BPMU(strPin, dblForceVoltage, dblClampCurrent)
        
                '========== Set Pattern ===============================
                TheHdw.WAIT dblWait
                
                '========== Measure Current ===============================
                Call MeasureI_BPMU(strPin, TempValue, lAve)
                                
                For i = 0 To Temp_BoardCnt - 1
                        tempValueSum = tempValueSum + TempValue(myboard_i)
                        myboard_i = myboard_i + Temp_ChanbyBoard
                Next i
                
                For site = 0 To nSite
                    If TheExec.sites.site(site).Active = True Then
                            retResult2(site) = tempValueSum
                    End If
                Next site
                
            Else
                '========== Set Condition ===============================
                Call TheCondition.SetCondition(strPowerCondition) '2
                '========== Force Voltage ===============================
                Call SetFVMI(strPin, dblForceVoltage, dblClampCurrent)
                '========== Set Pattern ===============================
                TheHdw.WAIT dblWait
                '========== Measure Current ===============================
                Call MeasureI(strPin, retResult2, lAve)
            End If
         
            
        'パターン停止
        
        Call DisconnectPins(strPin, ALL_SITE)
        'All_Open及びDisconnect自動生成化
        Call PowerDownAndDisconnect
        
    Else
    
    
        Call TheCondition.SetCondition(strSetParamCondition) '1
        lAve = GetSubCurrentAverageCount(GetInstanceName)
        dblClampCurrent = GetSubCurrentClampCurrent(GetInstanceName)
        dblWait = GetSubCurrentWaitTime(GetInstanceName)
        dblPinResourceName = GetSubCurrentPinResourceName(GetInstanceName)
        
        '========== Set Condition ===============================
        Call TheCondition.SetCondition(strPowerCondition) '2
        '========== Force Voltage ===============================
        Call SetFVMI(strPin, dblForceVoltage, dblClampCurrent)
        '========== Set Pattern ===============================
        TheHdw.WAIT dblWait
        '========== Measure Current ===============================
        Call MeasureI(strPin, retResult2, lAve)
            
        'パターン停止
        Call DisconnectPins(strPin, ALL_SITE)
        Call PowerDownAndDisconnect
        
    End If

    Call updateResult(strResultKey, retResult2)

    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob関数
    Call TheResult.Add(strResultKey, retResult)

End Function


Private Function GetChanType(ByVal PinList As String) As chtype

    Dim chanType As chtype
    Dim pinArr() As String
    Dim pinNum As Long
    Dim i As Long

    With TheExec.DataManager
        chanType = .chanType(PinList)

        If chanType = chUnk Then
            On Error GoTo INVALID_PINLIST
            Call .DecomposePinList(PinList, pinArr, pinNum)
            On Error GoTo -1

            chanType = .chanType(pinArr(0))
            For i = 0 To pinNum - 1
                If chanType <> .chanType(pinArr(i)) Then
                    chanType = chUnk
                    Exit For
                End If
            Next i
        End If
    End With

    GetChanType = chanType
    Exit Function

INVALID_PINLIST:
    GetChanType = chUnk

End Function


'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Serial_Test用:End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
Private Function getParam_BPMU_Parallel( _
    ByRef strPin As String, _
    ByRef dblForceVoltage As Double, _
    ByRef strSetParamCondition As String, _
    ByRef strPowerCondition As String, _
    ByRef strPatternCondition As String _
    ) As Boolean

    '変数取り込み
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_BPMU_PARA_ARGS) Then
        getParam_BPMU_Parallel = False
        Exit Function
    End If
    
    Dim tempstr As String      '一時保存変数
    Dim tempArrstr() As String '一時保存配列
    
On Error GoTo ErrHandler
    strPin = ArgArr(0)                      'Arg20: Test pin name
    dblForceVoltage = CDbl(ArgArr(1))       'Arg21: Force voltage (PPS value) for test pin.
    tempstr = ArgArr(2)
    tempArrstr = Split(tempstr, ",")
    strSetParamCondition = tempArrstr(0)        'Arg22-1: [Test Condition]'s condition name for common environment setup.
    strPowerCondition = tempArrstr(1)           'Arg22-2: [Test Condition]'s condition name for device setup.
    strPatternCondition = tempArrstr(2)         'Arg22-3: [Test Condition]'s condition name for patttern burst.
'    strResultKey = ArgArr(3)
On Error GoTo 0

    getParam_BPMU_Parallel = True
    Exit Function
    
ErrHandler:

    getParam_BPMU_Parallel = False
    Exit Function

End Function



'♪
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'subCurrent_Serial_Test用:End
'＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝＝
'♪
Private Function getParam_BPMU_Parallel_NonPattern( _
    ByRef strPin As String, _
    ByRef dblForceVoltage As Double, _
    ByRef strSetParamCondition As String, _
    ByRef strPowerCondition As String _
    ) As Boolean

    '変数取り込み
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_BPMU_PARA_ARGS) Then
        getParam_BPMU_Parallel_NonPattern = False
        Exit Function
    End If
    
    Dim tempstr As String      '一時保存変数
    Dim tempArrstr() As String '一時保存配列
    
On Error GoTo ErrHandler
    strPin = ArgArr(0)                      'Arg20: Test pin name
    dblForceVoltage = CDbl(ArgArr(1))       'Arg21: Force voltage (PPS value) for test pin.
    tempstr = ArgArr(2)
    tempArrstr = Split(tempstr, ",")
    strSetParamCondition = tempArrstr(0)        'Arg22-1: [Test Condition]'s condition name for common environment setup.
    strPowerCondition = tempArrstr(1)           'Arg22-2: [Test Condition]'s condition name for device setup.
'    strResultKey = ArgArr(3)
On Error GoTo 0

    getParam_BPMU_Parallel_NonPattern = True
    Exit Function
    
ErrHandler:

    getParam_BPMU_Parallel_NonPattern = False
    Exit Function

End Function

Private Function Sub_NonScenarioParaJudge_GetParameter( _
    ByRef dblLoJudgeLimit As Double, _
    ByRef dblHiJudgeLimit As Double, _
    ByRef dblHiLoLimValid As Long _
    ) As Boolean

On Error GoTo ErrHandler
    '変数取り込み
    '想定数よりと違う場合エラーコード
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        Sub_NonScenarioParaJudge_GetParameter = False
        Exit Function
    End If
    
'    strResultKey = ArgArr(0)
    
    If ArgArr(0) <> "" Then
        dblLoJudgeLimit = ArgArr(0)
    Else
        ArgArr(0) = 0
    End If
    
    If ArgArr(1) <> "" Then
        dblHiJudgeLimit = ArgArr(1)
    Else
        ArgArr(1) = 0
    End If
    
    dblHiLoLimValid = ArgArr(2)
    On Error GoTo 0

    Sub_NonScenarioParaJudge_GetParameter = True
    Exit Function
    
ErrHandler:

    Sub_NonScenarioParaJudge_GetParameter = False
    Exit Function

End Function

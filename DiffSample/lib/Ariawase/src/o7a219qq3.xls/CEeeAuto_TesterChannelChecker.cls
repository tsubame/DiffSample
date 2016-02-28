VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CEeeAuto_TesterChannelChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'概要:
'   コンディションチェッカ　電源情報取得モジュール
'
'目的:
'   そのときの条件設定を装置から取得する
'   スナップショットのまるパクリ
'作成者:
'   丸山
'   有川 編集 Ver1.1 2013/02/25

Option Explicit

'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const HDVIS_USE = 0      'HVDISリソースの使用　 0：未使用、0以外：使用  <PALSとEeeAuto両方で使用>

'SnapShotヘッダ情報用構造体
Private Type type_SS_HEADER
    tResourceName() As String
    tIdLabel() As String
    tPinName() As String
    tSiteNumber() As Long
    tChannelNumber() As Long
End Type

'APMU情報取得用構造体
Private Type Type_APMU
    tGate As Long
    tRelay As Long
    tLowPassFilter As Long
    tExternalSense As Long
    tAlarm As Long
    tMode As ApmuMode
    tClampValue As Double
    tForceValue As Double
    tIRange As ApmuIRange
    tVRange As ApmuVRange
    tGangPinFlag As Long
'    tMeasureResult As Double
End Type

'APMU取得時使用TMP構造体
Private Type Type_TMP_APMU
    tGate() As Long
    tRelay() As Long
    tLowPassFilter() As Long
    tExternalSense() As Long
    tAlarm() As Long
    tMode() As ApmuMode
    tClampValue() As Double
    tForceValue() As Double
    tIRange() As ApmuIRange
    tVRange() As ApmuVRange
    tGangPinFlag() As Long
'    tMeasureResult() As Double
End Type

'DPS情報取得用構造体
Private Type type_DPS
    tCurrentLimit As Double
    tPrimaryVoltage As Double
    tAlternateVoltage As Double
    tCurrentRange As DpsIRange
    tOutputSource As Long
    tForceRelay As String
    tSenseRelay As String
'    tMeasureResult As Double
    tGangPinFlag As Long
    tMeasureSamples As Long
End Type

'PPMU情報取得用構造体
Private Type type_PPMU_INFO
    tForceVoltage As Double
    tForceCurrent As Double
    tCurrentRange As Long
    tHighLimit As Double
    tLowLimit As Double
    tConnect As Boolean
    tForceType As String
'    tMeasureResult As Double
    tMeasureSamples As Long
End Type

'BPMU情報取得用構造体
Private Type type_BPMU_INFO
    tClampCurrent As Double
    tClampVoltage As Double
    tForceCurrent As Double
    tForceVoltage As Double
    tVoltageRange As Long
    tCurrentRange As Long
    tHighLimit As Double
    tLowLimit As Double
    tVoltmeterMode As Boolean
    tBpmuGate As Boolean
    tConnectDut As Boolean
    tForcingMode As String
    tMeasureMode As String
'    tMeasureResult As Double
End Type

'Digtal-ch(PE)情報取得用
Private Type type_PE_INFO
    tVDriveLo As Double
    tVDriveHi As Double
    tVCompareLo As Double
    tVCompareHi As Double
    tVClampLo As Double
    tVClampHi As Double
    tVThreshold As Double
    tISource As Double
    tISink As Double
    tPeConnect As Boolean
    tHvConnect As Boolean
    tPpmuConnect As Boolean
    tBpmuConnect As Boolean
    tD0 As Double
    tD1 As Double
    tD2 As Double
    tD3 As Double
    tR0 As Double
    tR1 As Double
    tHvVph As Double
    tHvIph As Double
    tHvTpr As Double
End Type

'HDVIS情報取得用
Private Type type_HDVIS
    tGate As Long
    tRelay As Long
    tLowPassFilter As Long
    tAlarmOpenDgs As Long
    tAlarmOverLoad As Long
    tMargePinFlag As Long
    tMode As Long
    tVRange As Long
    tIRange As Long
    tSlewRate As Long
    tRelayMode As Long
    tClampValue As Double
    tForceValue As Double
    tMeasureResult As Double
    tExtMode As Long
    tExtSendRelay As Long
    tExtTriggerRelay As Long
    tSetupEnable As Boolean
End Type

'APMU情報用構造体
Private Type type_APMU_INFO
    tSsHeader As type_SS_HEADER
    tApmuinf() As Type_APMU
End Type

'HDVIS情報用構造体
Private Type type_HDVIS_INFO
    tSsHeader As type_SS_HEADER
    tHdvisInf() As type_HDVIS
End Type

'DPS情報用構造体
Private Type type_DPS_INFO
    tSsHeader As type_SS_HEADER
    tDpsinf() As type_DPS
End Type

'I/O Pin(デジタルピン) 情報用構造体
Private Type type_IO_INFO
    tSsHeader As type_SS_HEADER
    tPeinf() As type_PE_INFO
    tPpmuinf() As type_PPMU_INFO
    tBpmuinf() As type_BPMU_INFO
End Type

'このモジュール独自の構造体
Private Type type_ALL_Resouces
    tApmuInfo As type_APMU_INFO
    tDchanInfo As type_IO_INFO
    tDpsInfo As type_DPS_INFO
    tHdvisInfo As type_HDVIS_INFO
End Type
Implements IEeeAuto_TestConditionChecker

'モードの定義
Private m_strMode As String

Private m_TesterConditions(1) As type_ALL_Resouces

Private Sub Class_Initialize()
    
    m_strMode = ""

End Sub

Private Sub IEeeAuto_TestConditionChecker_CheckBeforeCondition()

'内容:
'   電源参照テーブルの初期化
'
'注意事項:

    m_TesterConditions(0) = GetTesterResouceConditionInfo

End Sub

Private Function IEeeAuto_TestConditionChecker_CheckAfterCondition() As Boolean

    'DefaultをValid
    IEeeAuto_TestConditionChecker_CheckAfterCondition = True
    
    m_TesterConditions(1) = GetTesterResouceConditionInfo
        
    If (IsEqualTwoCondition(m_TesterConditions(0), m_TesterConditions(1))) Then
        '条件が一緒だった場合 InValidとする
        IEeeAuto_TestConditionChecker_CheckAfterCondition = False
    End If

End Function

Private Sub IEeeAuto_TestConditionChecker_SetOperationMode(ByVal strMode As String)
    
    m_strMode = strMode
    
End Sub

Private Sub IEeeAuto_TestConditionChecker_SetEndPosition(ByVal strPos As String)

    '最終位置はこのクラスにはない

End Sub
'=================================================================


Private Function GetTesterResouceConditionInfo() As type_ALL_Resouces
'内容:
'   テスタから応答を取得したそのものをかえす
'
'注意事項:
'   とくになし

    Dim tPinNameArr() As String
    Dim tChansArr() As Long
    Dim resourceChk As Boolean
    Dim ResourceName As String

    Dim tApmuinf As type_APMU_INFO 'APMU情報用構造体
    Dim tDpsinf As type_DPS_INFO 'DPS情報用構造体
    Dim tPeinf As type_IO_INFO 'PE情報用構造体
        
    '=====APMU======================================
    'APMUリソースを使用しているChannelを調べる
    resourceChk = mf_ChkResourcePin_Site0(chAPMU, tChansArr, tPinNameArr)
    ResourceName = "[APMU]" 'IP750リソース識別用ラベル[APMU]
                                                                                                
    'APMUリソースを使用していたら値をとる
    If resourceChk = True Then
        'APMU情報をテラダインAPIを使用して取得
        Call mf_Make_SsHeader_Site0(ResourceName, "*", tPinNameArr, tChansArr, tApmuinf.tSsHeader)
        Call mf_GetApmuInfo(tApmuinf.tSsHeader.tChannelNumber, tApmuinf)
        GetTesterResouceConditionInfo.tApmuInfo = tApmuinf
    End If
    Erase tPinNameArr
    Erase tChansArr
    
    
    '=====DPS======================================
    'DPSリソースを使用しているChannelを調べる
    resourceChk = mf_ChkResourcePin_Site0(chDPS, tChansArr, tPinNameArr)
    ResourceName = "[DPS]" 'IP750リソース識別用ラベル[DPS]
                                                                                                
    'DPSリソースを使用していたら値をとる
    If resourceChk = True Then
        'DPS情報のヘッダ部の情報を作成
        Call mf_Make_SsHeader_Site0(ResourceName, "*", tPinNameArr, tChansArr, tDpsinf.tSsHeader)
        'DPS情報をTERADYNE-APIから取得
        Call mf_GetDpsInfo(tDpsinf.tSsHeader.tChannelNumber, tDpsinf)
        GetTesterResouceConditionInfo.tDpsInfo = tDpsinf
    End If
    Erase tPinNameArr
    Erase tChansArr
    
    
    'DChanを見ないときは無視
    If m_strMode <> EEE_AUTO_TESTER_CHECKER_WITHOUT_DCHAN Then
    
        '=====Dchan(IO)======================================
        'I/O(PE)リソースを使用しているChannelを調べる
        resourceChk = mf_ChkResourcePin_Site0(chIO, tChansArr, tPinNameArr)
        ResourceName = "[PE]" 'IP750リソース識別用ラベル[PE]
    
        'I/O(PE)リソースを使用していたら値をとる
        If resourceChk = True Then
            'PE情報のヘッダ部の情報を作成
            Call mf_Make_SsHeader_Site0(ResourceName, "*", tPinNameArr, tChansArr, tPeinf.tSsHeader)
            'PE情報をTERADYNE-APIから取得
            Call mf_GetPeInfo(tPeinf.tSsHeader.tChannelNumber, tPeinf.tPeinf)
            'PPMU情報をTERADYNE-APIから取得
            Call mf_GetPpmuInfo(tPeinf.tSsHeader.tChannelNumber, tPeinf.tPpmuinf)
            'BPMU情報をTERADYNE-APIから取得
            Call mf_GetBpmuInfo(tPeinf.tSsHeader.tChannelNumber, tPeinf.tSsHeader.tPinName, tPeinf.tSsHeader.tSiteNumber, tPeinf.tBpmuinf)
            GetTesterResouceConditionInfo.tDchanInfo = tPeinf
        End If
        Erase tPinNameArr
        Erase tChansArr
        
    End If
   
#If HDVIS_USE <> 0 Then

    Const CH_HDVIS = 36
    ResourceName = "[HDVIS]" 'IP750リソース識別用ラベル[HDVIS]

    '=====HDVIS======================================
    'HDVISリソースを使用しているChannelを調べる
    resourceChk = mf_ChkResourcePin(CH_HDVIS, tChansArr, tPinNameArr)

    'HDVISリソースを使用していないときは終了
    If resourceChk = False Then
        'HDVIS情報のヘッダ部の情報を作成
        Call mf_Make_SsHeader(ResourceName, testIdLabel, tPinNameArr, tChansArr, tHdvisInf.tSsHeader)
    End If
    MsgBox "No Debug!! Don't Use HDVIS Condition Auto omission"
#End If
                                                                                                     
End Function

Private Function IsEqualTwoCondition(ByRef RHS As type_ALL_Resouces, ByRef lhs As type_ALL_Resouces) As Boolean
'内容:
'   type_ALL_Resouces構造体の operator =
'
'注意事項:
'   比較内容はひとまずこれで。

    IsEqualTwoCondition = False
    
    Dim i As Long

    For i = 0 To UBound(RHS.tApmuInfo.tApmuinf)
        If RHS.tApmuInfo.tApmuinf(i).tClampValue <> lhs.tApmuInfo.tApmuinf(i).tClampValue Then
            Exit Function
        End If
        If RHS.tApmuInfo.tApmuinf(i).tForceValue <> lhs.tApmuInfo.tApmuinf(i).tForceValue Then
            Exit Function
        End If
        If RHS.tApmuInfo.tApmuinf(i).tIRange <> lhs.tApmuInfo.tApmuinf(i).tIRange Then
            Exit Function
        End If
        If RHS.tApmuInfo.tApmuinf(i).tGangPinFlag <> lhs.tApmuInfo.tApmuinf(i).tGangPinFlag Then
            Exit Function
        End If
        If RHS.tApmuInfo.tApmuinf(i).tRelay <> lhs.tApmuInfo.tApmuinf(i).tRelay Then
            Exit Function
        End If
    Next i

    For i = 0 To UBound(RHS.tDpsInfo.tDpsinf)
        If RHS.tDpsInfo.tDpsinf(i).tPrimaryVoltage <> lhs.tDpsInfo.tDpsinf(i).tPrimaryVoltage Then
            Exit Function
        End If
        If RHS.tDpsInfo.tDpsinf(i).tForceRelay <> lhs.tDpsInfo.tDpsinf(i).tForceRelay Then
            Exit Function
        End If
    Next i

    'DChanを見ないときは無視
    If m_strMode <> EEE_AUTO_TESTER_CHECKER_WITHOUT_DCHAN Then
        For i = 0 To UBound(RHS.tDchanInfo.tBpmuinf)
            With RHS.tDchanInfo.tBpmuinf(i)
                If .tBpmuGate <> lhs.tDchanInfo.tBpmuinf(i).tBpmuGate Then
                    Exit Function
                End If
            End With
        Next
        
        For i = 0 To UBound(RHS.tDchanInfo.tPeinf)
            With RHS.tDchanInfo.tPeinf(i)
                If .tPpmuConnect <> lhs.tDchanInfo.tPeinf(i).tPpmuConnect Then
                    Exit Function
                End If
                If .tBpmuConnect <> lhs.tDchanInfo.tPeinf(i).tBpmuConnect Then
                    Exit Function
                End If
                If .tPeConnect <> lhs.tDchanInfo.tPeinf(i).tPeConnect Then
                    Exit Function
                End If
                If .tVDriveHi <> lhs.tDchanInfo.tPeinf(i).tVDriveHi Then
                    Exit Function
                End If
                If .tVDriveLo <> lhs.tDchanInfo.tPeinf(i).tVDriveLo Then
                    Exit Function
                End If
            End With
        Next
        
        For i = 0 To UBound(RHS.tDchanInfo.tPpmuinf)
            With RHS.tDchanInfo.tPpmuinf(i)
                If .tForceType <> lhs.tDchanInfo.tPpmuinf(i).tForceType Then
                    Exit Function
                End If
                If .tForceVoltage <> lhs.tDchanInfo.tPpmuinf(i).tForceVoltage Then
                    Exit Function
                End If
                If .tForceCurrent <> lhs.tDchanInfo.tPpmuinf(i).tForceCurrent Then
                    Exit Function
                End If
                If .tCurrentRange <> lhs.tDchanInfo.tPpmuinf(i).tCurrentRange Then
                    Exit Function
                End If
            End With
        Next
    End If

    IsEqualTwoCondition = True '最後まできたらおめでとう
    
End Function


'指定リソースを使用しているチャンネルとPinNameを調べる Site0のみ版
Private Function mf_ChkResourcePin_Site0(ByVal ResourceName As chtype, _
ByRef rChansArr() As Long, _
ByRef rPinNameArr() As String) As Boolean

    Dim rPinCnt As Long
    Dim rChCnt As Long
    Dim rSiteCnt As Long
    Dim rAllPinsStr As String
    Dim funcName As String
    
    funcName = "@mf_ChkResourcePin"

    '指定リソースを使用しているPIN情報を取得
    Call TheExec.DataManager.GetPinNames(rPinNameArr, ResourceName, rPinCnt)
                                                   
    '指定されたりソースが、定義されていないときはFalseを返して終了
    If rPinCnt = 0 Then
        mf_ChkResourcePin_Site0 = False
        Exit Function
    End If
                                                
    '指定リソースとして定義されているすべてのPINの名前をカンマ区切りで作成　　("P_PIN1,P_PIN2, .....")
    rAllPinsStr = mf_Make_PinNameStr(rPinNameArr)
                    
    '指定リソースとして定義されているすべてのPINのチャンネル番号を取得
    Call TheExec.DataManager.GetChanList(rAllPinsStr, 0, ResourceName, _
    rChansArr, rChCnt, rSiteCnt, "Resource Pin Check Error" & funcName)

    mf_ChkResourcePin_Site0 = True

End Function

'スナップショットのヘッダ部を作成
Private Sub mf_Make_SsHeader_Site0(ByVal ResourceName As String, ByVal tstIdLabel As String, _
ByRef pinNameArr() As String, ByRef chansArr() As Long, ByRef ssHeaderInf As type_SS_HEADER)
    
    Dim PinCnt As Long
    Dim siteNumCnt As Long
    Dim lopCnt As Long

    'ヘッダ用の構造体の箱を準備
    With ssHeaderInf
        ReDim .tResourceName(UBound(chansArr))
        ReDim .tIdLabel(UBound(chansArr))
        ReDim .tPinName(UBound(chansArr))
        ReDim .tSiteNumber(UBound(chansArr))
    End With
    
    'チャンネル番号配列をもらう
    ssHeaderInf.tChannelNumber = chansArr
    
    lopCnt = 0
    
    'サイトNOとピン名、リソース名称、ラベル情報の作成
    For PinCnt = 0 To UBound(pinNameArr) '対象ピン LOOP
            
        With ssHeaderInf
            .tResourceName(lopCnt) = ResourceName
            .tIdLabel(lopCnt) = tstIdLabel
            .tPinName(lopCnt) = pinNameArr(PinCnt)
            .tSiteNumber(lopCnt) = 0 'Site0固定
        End With
                   
        lopCnt = lopCnt + 1
        
    Next PinCnt

End Sub



'=================================================================
'ここより下はXLibSnapshotIP750Mod そのまま
'=================================================================

'指定リソースを使用しているチャンネルとPinNameを調べる
Private Function mf_ChkResourcePin(ByVal ResourceName As chtype, _
ByRef rChansArr() As Long, _
ByRef rPinNameArr() As String) As Boolean

    Dim rPinCnt As Long
    Dim rChCnt As Long
    Dim rSiteCnt As Long
    Dim rAllPinsStr As String
    Dim funcName As String
    
    funcName = "@mf_ChkResourcePin"

    '指定リソースを使用しているPIN情報を取得
    Call TheExec.DataManager.GetPinNames(rPinNameArr, ResourceName, rPinCnt)
                                                   
    '指定されたりソースが、定義されていないときはFalseを返して終了
    If rPinCnt = 0 Then
        mf_ChkResourcePin = False
        Exit Function
    End If
                                                
    '指定リソースとして定義されているすべてのPINの名前をカンマ区切りで作成　　("P_PIN1,P_PIN2, .....")
    rAllPinsStr = mf_Make_PinNameStr(rPinNameArr)
                    
    '指定リソースとして定義されているすべてのPINのチャンネル番号を取得
    Call TheExec.DataManager.GetChanList(rAllPinsStr, -1, ResourceName, _
    rChansArr, rChCnt, rSiteCnt, "Resource Pin Check Error" & funcName)

    mf_ChkResourcePin = True

End Function

'もらった配列に格納されている要素の名前を、カンマ区切り形式で作成
Private Function mf_Make_PinNameStr(ByRef pinNameArr() As String) As String

    Dim tLoopCnt As Long
        
    '配列に格納されているすべてのPINの名前を、カンマ区切り形式で作成　　("P_PIN1,P_PIN2, .....")
    mf_Make_PinNameStr = pinNameArr(0)
    
    For tLoopCnt = 1 To UBound(pinNameArr)
        mf_Make_PinNameStr = mf_Make_PinNameStr & "," & pinNameArr(tLoopCnt)
    Next tLoopCnt

End Function

Private Sub mf_GetApmuInfo(ByRef apmuChans() As Long, ByRef typeApmuInf As type_APMU_INFO)
'内容:
'   指定CHのAPMU情報を取得します
'
'[apmuChans]      In  情報を取得したいAPMUのCH
'[typeApmuInf]    Out  取得結果格納用APMU構造体
'
    Call mf_MakeApmuInfo(apmuChans, typeApmuInf)

End Sub

Private Sub mf_GetDpsInfo(ByRef dpsChans() As Long, ByRef typeDpsInf As type_DPS_INFO)
'内容:
'   指定CHのDPS情報を取得します
'
'[dpsChans]       In  情報を取得したいDPSのCH
'[typeDpsInf]     Out  取得結果格納用DPS構造体
'
    Call mf_MakeDpsInfo(dpsChans, typeDpsInf)

End Sub

Private Sub mf_GetPeInfo(ByRef peChans() As Long, ByRef typePeInf() As type_PE_INFO)
'内容:
'   指定CHのPE情報を取得します
'
'[peChans]        In   情報を取得したいI/O(PE)のCH
'[typePeInf]      Out  取得結果格納用PE構造体
'
    Call mf_MakePeInfo(peChans, typePeInf)

End Sub

Private Sub mf_GetPpmuInfo(ByRef ppmuChans() As Long, ByRef typePpmuInf() As type_PPMU_INFO)
'内容:
'   指定CHのPPMU情報を取得します
'
'[ppmuChans]        In   情報を取得したいI/O(PPMU)のCH
'[typePpmuInf]      Out  取得結果格納用PPMU構造体
'
    Call mf_MakePpmuInfo(ppmuChans, typePpmuInf)

End Sub

Private Sub mf_GetBpmuInfo(ByRef bpmuChans() As Long, _
ByRef bpmuPins() As String, _
ByRef siteNum() As Long, _
ByRef typeBpmuInf() As type_BPMU_INFO)
'内容:
'   指定CHのPPMU情報を取得します
'
'[bpmuChans]        In   情報を取得したいI/O(BPMU)のCH
'[bpmuPins]         In   情報を取得したいI/O(BPMU)のピン名
'[siteNum]          In   情報を取得したいI/O(BPMU)のサイト番号
'[typeBpmuInf]      Out  取得結果格納用BPMU構造体
'
    Call mf_MakeBpmuInfo(bpmuChans, bpmuPins, siteNum, typeBpmuInf)

End Sub

Private Sub mf_GetHdvisInfo(ByRef hdvischans() As Long, ByRef typeHdvisInf As type_HDVIS_INFO)
'内容:
'   指定CHのHDVIS情報を取得します
'
'[hdvisChans]      In  情報を取得したいHDVISのCH
'[typeHdvisInf]    Out 取得結果格納用HDVIS構造体
'
    Call mf_MakeHdvisInfo(hdvischans, typeHdvisInf)

End Sub

'指定CHのAPMU情報をTERADYNE-APIから取得
Private Sub mf_MakeApmuInfo(ByRef apmuChans() As Long, ByRef typeApmuInf As type_APMU_INFO)
    
    Dim tmpApmuMode() As ApmuMode
    Dim tmpForceValue() As Double
    Dim tmpClampValue() As Double
    Dim myRetVrange() As ApmuVRange
    Dim myRetIrange() As ApmuIRange
    Dim tchCnt As Long
    Dim read_apmu As Boolean
    
    Dim tmpApmuInf As Type_TMP_APMU
        

    '指定CHのAPMUリソース状況を取得
    
    With tmpApmuInf
        .tGate = TheHdw.APMU.Chans(apmuChans).Gate
        .tRelay = TheHdw.APMU.Chans(apmuChans).relay
        .tLowPassFilter = TheHdw.APMU.Chans(apmuChans).LowPassFilter
        .tExternalSense = TheHdw.APMU.Chans(apmuChans).ExternalSense
        .tAlarm = TheHdw.APMU.Chans(apmuChans).alarm
    End With

    '指定CHのAPMUモード情報を取得。
    With tmpApmuInf
        TheHdw.APMU.Chans(apmuChans).ReadRangesAndMode .tMode, .tVRange, .tIRange
    End With
    
    '指定CHのAPMU情報取得用の構造体の箱を準備
    With tmpApmuInf
        ReDim .tForceValue(UBound(apmuChans))
        ReDim .tClampValue(UBound(apmuChans))
    End With
                
    '指定CHのギャング接続の確認と、メータの読み取り値の取得
    With tmpApmuInf
        .tGangPinFlag = TheHdw.APMU.Chans(apmuChans).GangedChannels     'ギャング接続状態の確認
'        Call TheHdw.APMU.chans(apmuChans).measure(1, .tMeasureResult)   'メータ読み取り値を取得
    End With
                                
    '指定CHのAPMUモード別判定とForce、Clampの値の取得
    'APMUスナップショット構造体に結果詰め込み
    ReDim typeApmuInf.tApmuinf(UBound(apmuChans))
    
    For tchCnt = 0 To UBound(apmuChans) '対象CH LOOP
    'APMUモードにあわせて、レンジと設定値を取得
        Select Case tmpApmuInf.tMode(tchCnt)
            Case apmuForceIMeasureV:
                read_apmu = TheHdw.APMU.Chans(apmuChans(tchCnt)).ReadForceCurrents(myRetIrange, tmpForceValue)
                read_apmu = TheHdw.APMU.Chans(apmuChans(tchCnt)).ReadClampVoltages(myRetVrange, tmpClampValue)
            Case apmuForceVMeasureI:
                read_apmu = TheHdw.APMU.Chans(apmuChans(tchCnt)).ReadForceVoltages(myRetVrange, tmpForceValue)
                read_apmu = TheHdw.APMU.Chans(apmuChans(tchCnt)).ReadClampCurrents(myRetIrange, tmpClampValue)
            Case apmuMeasureV:
                read_apmu = TheHdw.APMU.Chans(apmuChans(tchCnt)).ReadClampVoltages(myRetVrange, tmpClampValue)
                ReDim tmpForceValue(UBound(apmuChans))
            End Select
            
        'APMU情報構造体に、CH LOOPで取得結果詰め込み
        With typeApmuInf.tApmuinf(tchCnt)
            .tAlarm = tmpApmuInf.tAlarm(tchCnt)
            .tClampValue = tmpClampValue(0)
            .tExternalSense = tmpApmuInf.tExternalSense(tchCnt)
            .tForceValue = tmpForceValue(0)
            .tGangPinFlag = tmpApmuInf.tGangPinFlag(tchCnt)
            .tGate = tmpApmuInf.tGate(tchCnt)
            .tIRange = tmpApmuInf.tIRange(tchCnt)
            .tLowPassFilter = tmpApmuInf.tLowPassFilter(tchCnt)
'            .tMeasureResult = tmpApmuInf.tMeasureResult(tchCnt)
            .tMode = tmpApmuInf.tMode(tchCnt)
            .tRelay = tmpApmuInf.tRelay(tchCnt)
            .tVRange = tmpApmuInf.tVRange(tchCnt)
        End With
    
    Next tchCnt
    
End Sub

'指定CHのDPS情報をTERADYNE-APIから取得
Private Sub mf_MakeDpsInfo(ByRef dpsChans() As Long, ByRef typeDpsInf As type_DPS_INFO)
    
    Dim tchCnt As Long
    Dim tmpMesureVal() As Double
    Dim tmpCurrentLimit As Variant
    Dim tmpPrimaryVoltage As Variant
    Dim tmpAlternateVoltage As Variant
    Dim aveCnt As Long

    'DPS情報の箱を準備
    ReDim typeDpsInf.tDpsinf(UBound(dpsChans))
    
    For tchCnt = 0 To UBound(dpsChans) Step 1
        
        With typeDpsInf.tDpsinf(tchCnt)
            'リソース設定状態取得
            tmpCurrentLimit = TheHdw.DPS.Chans(dpsChans(tchCnt)).CurrentLimit
            .tCurrentLimit = tmpCurrentLimit(0)
            
            tmpPrimaryVoltage = TheHdw.DPS.Chans(dpsChans(tchCnt)).forceValue(dpsPrimaryVoltage)
            .tPrimaryVoltage = tmpPrimaryVoltage(0)
                        
            tmpAlternateVoltage = TheHdw.DPS.Chans(dpsChans(tchCnt)).forceValue(dpsAlternateVoltage)
            .tAlternateVoltage = tmpAlternateVoltage(0)
            
            .tCurrentRange = TheHdw.DPS.Chans(dpsChans(tchCnt)).CurrentRange
            .tOutputSource = TheHdw.DPS.Chans(dpsChans(tchCnt)).OutputSource

            'リレー接続状態取得
            If TheHdw.DPS.Chans(dpsChans(tchCnt)).ForceRelayClosed = True Then
                .tForceRelay = "Closed"
            Else
                .tForceRelay = "Open"
            End If

            If TheHdw.DPS.Chans(dpsChans(tchCnt)).SenseRelayClosed = True Then
                .tSenseRelay = "Closed"
            Else
                .tSenseRelay = "Open"
            End If
                
            'メーターのアベレージ数取得
            .tMeasureSamples = TheHdw.DPS.Samples
            
'            '電流計の電流値取得
'            Call TheHdw.DPS.chans(dpsChans(tchCnt)).MeasureCurrents(.tCurrentRange, tmpMesureVal)
'
'            .tMeasureResult = 0
'
'            For aveCnt = 0 To UBound(tmpMesureVal)
'                .tMeasureResult = .tMeasureResult + tmpMesureVal(aveCnt)
'            Next aveCnt
'
'            .tMeasureResult = .tMeasureResult / .tMeasureSamples
        
        End With
                                                              
    Next tchCnt
    
End Sub

'指定CHのPE情報をTERADYNE-APIから取得
Private Sub mf_MakePeInfo(ByRef peChans() As Long, ByRef typePeInf() As type_PE_INFO)
    
    Dim tchCnt As Long
    
    ReDim typePeInf(UBound(peChans))
    
    For tchCnt = 0 To UBound(peChans) Step 1
        
        With typePeInf(tchCnt)
            .tVDriveLo = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chVDriveLo)
            .tVDriveHi = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chVDriveHi)
            .tVClampLo = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chVCL)
            .tVClampHi = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chVCH)
            .tVCompareLo = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chVCompareLo)
            .tVCompareHi = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chVCompareHi)
            .tVThreshold = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chVT)
            .tISource = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chISource)
            .tISink = TheHdw.PinLevels.chan(peChans(tchCnt)).readPinLevels(chISink)
                                                                  
            'High Voltage Status
            If (peChans(tchCnt) Mod 32) = 0 Or ((peChans(tchCnt) - 4) Mod 32) = 0 Then
                Call TheHdw.PinLevels.chan(peChans(tchCnt)).ReadHighVoltageParams(.tHvVph, .tHvIph, .tHvTpr)
            End If
                        
'{
            'Rdge Setの設定値をいくらにすればよいのか判らないので封印
'            Call TheHdw.Digital.Timing.chan(peChans(tchCnt)).readEdgeTimingRAM(0)
'            .tD0 = TheHdw.Digital.Timing.EdgeTime(chEdgeD0)
'            .tD1 = TheHdw.Digital.Timing.EdgeTime(chEdgeD1)
'            .tD2 = TheHdw.Digital.Timing.EdgeTime(chEdgeD2)
'            .tD3 = TheHdw.Digital.Timing.EdgeTime(chEdgeD3)
'            .tR0 = TheHdw.Digital.Timing.EdgeTime(chEdgeR0)
'            .tR1 = TheHdw.Digital.Timing.EdgeTime(chEdgeR1)
'}
                                                                  
            'IOピン(HV)のリレー接続状態確認
             .tHvConnect = mf_ChkIoRelayStat(peChans(tchCnt), "HV")
            'IOピン(PE)のリレー接続状態確認
             .tPeConnect = mf_ChkIoRelayStat(peChans(tchCnt), "PE")
            'IOピン(PPMU)のリレー接続状態確認
             .tPpmuConnect = mf_ChkIoRelayStat(peChans(tchCnt), "PPMU")
            'IOピン(BPMU)のリレー接続状態確認
             .tBpmuConnect = mf_ChkIoRelayStat(peChans(tchCnt), "BPMU")
        
        End With

    Next tchCnt
    
End Sub

'指定CHのPPMU情報をTERADYNE-APIから取得
Private Sub mf_MakePpmuInfo(ByRef ppmuChans() As Long, ByRef typePpmuInf() As type_PPMU_INFO)
    Dim tchCnt As Long
    Dim tmpMeasureVal() As Double
    Dim aveCnt As Long
    
    ReDim typePpmuInf(UBound(ppmuChans))
    
    For tchCnt = 0 To UBound(ppmuChans) Step 1
        
        With typePpmuInf(tchCnt)
            .tCurrentRange = TheHdw.PPMU.Chans(ppmuChans(tchCnt)).CurrentRange
            .tForceVoltage = TheHdw.PPMU.Chans(ppmuChans(tchCnt)).ForceVoltage(.tCurrentRange)
            .tForceCurrent = TheHdw.PPMU.Chans(ppmuChans(tchCnt)).ForceCurrent(.tCurrentRange)
            .tHighLimit = TheHdw.PPMU.Chans(ppmuChans(tchCnt)).TestLimitHigh
            .tLowLimit = TheHdw.PPMU.Chans(ppmuChans(tchCnt)).TestLimitLow
                                       
            If TheHdw.PPMU.Chans(ppmuChans(tchCnt)).IsForcingVoltage <> True Then
                .tForceType = "AMPS"
'                Call TheHdw.PPMU.chans(ppmuChans(tchCnt)).MeasureVoltages(tmpMeasureVal)
            Else
                .tForceType = "VOLTS"
'                Call TheHdw.PPMU.chans(ppmuChans(tchCnt)).MeasureCurrents(tmpMeasureVal)
            End If
                                                                                          
'            .tMeasureResult = 0
'            .tMeasureSamples = UBound(tmpMeasureVal) + 1
'
'            For aveCnt = 0 To UBound(tmpMeasureVal)
'                .tMeasureResult = .tMeasureResult + tmpMeasureVal(aveCnt)
'            Next aveCnt
'
'            .tMeasureResult = .tMeasureResult / .tMeasureSamples
                                                                                                                                                            
            'IOピンのリレー接続状態確認
             .tConnect = mf_ChkIoRelayStat(ppmuChans(tchCnt), "PPMU")
                            
        End With

    Next tchCnt
    
End Sub

'指定CHのBPMU情報をTERADYNE-APIから取得
Private Sub mf_MakeBpmuInfo(ByRef bpmuChans() As Long, _
ByRef bpmuPins() As String, _
ByRef siteNum() As Long, _
ByRef typeBpmuInf() As type_BPMU_INFO)

    Dim tchCnt As Long
    Dim tmpIrange() As Long
    Dim tmpVrange() As Long
    Dim tmpFvMode As Boolean
    Dim tmpMvMode As Boolean
    Dim tmpMeasureVal() As Double

    ReDim typeBpmuInf(UBound(bpmuChans))
        
    For tchCnt = 0 To UBound(bpmuChans) Step 1
        
        Call TheHdw.BPMU.Chans(bpmuChans(tchCnt)).ReadDriverRanges(tmpIrange, tmpVrange)
        
        With typeBpmuInf(tchCnt)
            .tCurrentRange = tmpIrange(0)
            .tVoltageRange = tmpVrange(0)
            .tClampCurrent = TheHdw.BPMU.Chans(bpmuChans(tchCnt)).ClampCurrent(.tCurrentRange)
            .tClampVoltage = TheHdw.BPMU.Chans(bpmuChans(tchCnt)).ClampVoltage(.tVoltageRange)
            .tForceCurrent = TheHdw.BPMU.Chans(bpmuChans(tchCnt)).ForceCurrent(.tCurrentRange)
            .tForceVoltage = TheHdw.BPMU.Chans(bpmuChans(tchCnt)).ForceVoltage(.tVoltageRange)
            .tHighLimit = TheHdw.BPMU.Chans(bpmuChans(tchCnt)).TestLimitHigh
            .tLowLimit = TheHdw.BPMU.Chans(bpmuChans(tchCnt)).TestLimitLow
            .tBpmuGate = TheHdw.BPMU.Chans(bpmuChans(tchCnt)).GateOn
            .tConnectDut = mf_ChkIoRelayStat(bpmuChans(tchCnt), "BPMU")
        End With
                                                                                                                                    
        '電圧印加、電流印加のモード
        tmpFvMode = TheHdw.BPMU.Pins(bpmuPins(tchCnt)).BpmuIsForcingVoltage(siteNum(tchCnt))
        '電圧測定、電流測定のモード
        tmpMvMode = TheHdw.BPMU.Pins(bpmuPins(tchCnt)).BpmuIsMeasuringVoltage(siteNum(tchCnt))
        
        With typeBpmuInf(tchCnt)
            If tmpFvMode = True Then
                .tForcingMode = "FV"
            Else
                .tForcingMode = "FI"
            End If
                                                                          
            If tmpMvMode = True Then
                .tMeasureMode = "MV"
            Else
                .tMeasureMode = "MI"
            End If
        End With
        
        'メーターリード
'        Call TheHdw.BPMU.chans(bpmuChans(tchCnt)).measure(1, tmpMeasureVal)
'        typeBpmuInf(tchCnt).tMeasureResult = tmpMeasureVal(0)
    
    Next tchCnt
    
End Sub

'指定CHのHDVIS情報をTERADYNE-APIから取得
Private Sub mf_MakeHdvisInfo(ByRef hdvischans() As Long, ByRef typeHdvisInf As type_HDVIS_INFO)
    
    Dim tmpGateStat() As Long
    Dim tmpRelayStat() As Long
    Dim tmpLowPassFilter() As Long
    Dim tmpAlarmOpnDgs() As Long
    Dim tmpAlarmOverLoad() As Long
    Dim tmpMergeFlg() As Long
    Dim tmpHdvisMode() As Long
    Dim tmpVrange() As Long
    Dim tmpIrange() As Long
    Dim tmpSlewRate() As Long
    Dim tmpRelayMode As Long
    Dim tmpForceValue() As Double
    Dim tmpClampValue() As Double
    Dim tmpMeasureValue() As Double
    Dim tmpExtMode() As Long
    Dim tmpExtSendRelay As Long
    Dim tmpExtTriggerRelay As Long
    Dim myForceValue() As Double
    Dim myClampValue() As Double
    Dim myRetIrange() As Long
    Dim myRetVrange() As Long
    Dim tchCnt As Long
    Dim readHdvis As Boolean
    Dim hdvisBoardNo As Long
    Dim setupEnable As Boolean
    
    'HDVISをサポートしていないIG-XLでコンパイルエラーとなるのを
    '回避するためTheHdw.HDVISを置き換え
'    Dim myHdvis As HdwDrivers.DriverHDVIS
    Dim myHdvis As Object
    Set myHdvis = TheHdw.HDVIS
                       
    'HDVIS情報用構造体準備
    ReDim typeHdvisInf.tHdvisInf(UBound(hdvischans))
            
    'CH毎、HDVISパラメータ状態取得
    With myHdvis.Chans(hdvischans)
        tmpGateStat = .Gate
        tmpRelayStat = .relay
        tmpLowPassFilter = .LowPassFilter
        tmpAlarmOpnDgs = .alarm(0)    'hdvisAlarmOpenDGS=0
        tmpAlarmOverLoad = .alarm(1)  'hdvisAlarmOverLoad=1
        tmpMergeFlg = .MergedChannels
        Call .ReadExternalModes(tmpExtMode)
        Call .ReadSlewRates(tmpSlewRate)
        Call .ReadRangesAndMode(tmpHdvisMode, tmpVrange, tmpIrange)
    End With
        
    'リレーモード取得（リレーモードはすべてのCH共通、CH毎の設定はなし）
    tmpRelayMode = myHdvis.RelayMode
    
    'Measure値取得
    Call myHdvis.Chans(hdvischans).measure(1, tmpMeasureValue)
                
    '指定CHのHDVISモード別判定とForce、Clampの値の取得
    For tchCnt = 0 To UBound(hdvischans)  '取得対象CH LOOP
        
        'Forceモードに応じてレンジと設定値を取得
        Select Case tmpHdvisMode(tchCnt)
            Case 1 'hdvisForceIMeasureV: 'FI
                readHdvis = myHdvis.Chans(hdvischans(tchCnt)).ReadForceCurrents(myRetIrange, myForceValue)
                readHdvis = myHdvis.Chans(hdvischans(tchCnt)).ReadClampVoltages(myRetVrange, myClampValue)
            Case 0 'hdvisForceVMeasureI: 'FV
                readHdvis = myHdvis.Chans(hdvischans(tchCnt)).ReadForceVoltages(myRetVrange, myForceValue)
                readHdvis = myHdvis.Chans(hdvischans(tchCnt)).ReadClampCurrents(myRetIrange, myClampValue)
            Case 4 'hdvisMeasureV: 'MV HDVISはMVモード時にV-Clampの機能は無し
                ReDim myForceValue(0)
                myForceValue(0) = 0#
                ReDim myClampValue(0)
                myClampValue(0) = 0#
        End Select
        
        'ボード単位で存在する設定の状態取得
        With myHdvis
            hdvisBoardNo = .SlotNumber(hdvischans(tchCnt)) 'CH番号からボード番号取得
            setupEnable = .board(hdvisBoardNo).Setup.Enable
            tmpExtSendRelay = .board(hdvisBoardNo).ExternalSend.relay       '設定はボード毎
            tmpExtTriggerRelay = .board(hdvisBoardNo).ExternalTrigger.relay '設定はボード毎
        End With
        
        'HDVIS情報、構造体へ取得結果を詰め込み
        With typeHdvisInf.tHdvisInf(tchCnt)
            .tGate = tmpGateStat(tchCnt)
            .tRelay = tmpRelayStat(tchCnt)
            .tLowPassFilter = tmpLowPassFilter(tchCnt)
            .tAlarmOpenDgs = tmpAlarmOpnDgs(tchCnt)
            .tAlarmOverLoad = tmpAlarmOverLoad(tchCnt)
            .tMargePinFlag = tmpMergeFlg(tchCnt)
            .tMode = tmpHdvisMode(tchCnt)
            .tVRange = tmpVrange(tchCnt)
            .tIRange = tmpIrange(tchCnt)
            .tSlewRate = tmpSlewRate(tchCnt)
            .tRelayMode = tmpRelayMode
            .tForceValue = myForceValue(0)
            .tClampValue = myClampValue(0)
            .tMeasureResult = tmpMeasureValue(tchCnt)
            .tExtMode = tmpExtMode(tchCnt)
            .tExtSendRelay = tmpExtSendRelay
            .tExtTriggerRelay = tmpExtTriggerRelay
            .tSetupEnable = setupEnable
        End With
    
    Next tchCnt
    
    Set myHdvis = Nothing
    
End Sub

'デジタルPINのリレー状態を確認し、指定リソースの接続状態を返します。
Private Function mf_ChkIoRelayStat(DigitalChNo As Long, ChkResourceName As String) As Boolean

    Select Case mf_GetIoRelayStat(DigitalChNo)
        
        Case rlyPE:
            If ChkResourceName = "PE" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
        
        Case rlyPPMU:
            If ChkResourceName = "PPMU" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
        
        Case rlyBPMU:
            If ChkResourceName = "BPMU" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
        
        Case rlyHV:
            If ChkResourceName = "HV" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
        
        Case rlyDisconnect:
            mf_ChkIoRelayStat = False
        
        Case rlyPPMU_PE:
            If ChkResourceName = "PPMU" Or ChkResourceName = "PE" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
                
        Case rlyPPMU_HV:
            If ChkResourceName = "PPMU" Or ChkResourceName = "HV" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
        
        Case rlyBPMU_PE:
            If ChkResourceName = "BPMU" Or ChkResourceName = "PE" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
                
        Case rlyBPMU_PPMU:
            If ChkResourceName = "BPMU" Or ChkResourceName = "PPMU" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
                
        Case rlyBPMU_HV:
            If ChkResourceName = "BPMU" Or ChkResourceName = "HV" Then
                mf_ChkIoRelayStat = True
            Else
                mf_ChkIoRelayStat = False
            End If
                
    End Select

End Function
'デジタルPinのリレー接続状態を取得する
Private Function mf_GetIoRelayStat(chNumber As Long) As RlyType
    Dim rlyStat As RlyType

    On Error GoTo RLY_DISCON

    mf_GetIoRelayStat = TheHdw.Digital.relays.chan(chNumber).whichChanRelay
    
    Exit Function

RLY_DISCON:
    mf_GetIoRelayStat = rlyDisconnect

End Function



'スナップショットのヘッダ部を作成
Private Sub mf_Make_SsHeader(ByVal ResourceName As String, ByVal tstIdLabel As String, _
ByRef pinNameArr() As String, ByRef chansArr() As Long, ByRef ssHeaderInf As type_SS_HEADER)
    
    Dim PinCnt As Long
    Dim siteNumCnt As Long
    Dim lopCnt As Long

    'ヘッダ用の構造体の箱を準備
    With ssHeaderInf
        ReDim .tResourceName(UBound(chansArr))
        ReDim .tIdLabel(UBound(chansArr))
        ReDim .tPinName(UBound(chansArr))
        ReDim .tSiteNumber(UBound(chansArr))
    End With
    
    'チャンネル番号配列をもらう
    ssHeaderInf.tChannelNumber = chansArr
    
    lopCnt = 0
    
    'サイトNOとピン名、リソース名称、ラベル情報の作成
    For PinCnt = 0 To UBound(pinNameArr) '対象ピン LOOP
        For siteNumCnt = 0 To TheExec.sites.ExistingCount - 1 'マルチサイト LOOP
            
            With ssHeaderInf
                .tResourceName(lopCnt) = ResourceName
                .tIdLabel(lopCnt) = tstIdLabel
                .tPinName(lopCnt) = pinNameArr(PinCnt)
                .tSiteNumber(lopCnt) = siteNumCnt
            End With
                       
            lopCnt = lopCnt + 1
        Next siteNumCnt
    Next PinCnt

End Sub


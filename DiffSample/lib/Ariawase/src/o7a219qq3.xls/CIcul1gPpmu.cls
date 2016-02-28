VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CIcul1gPpmu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@@@IP750EX Support
'ICUL1G のPPMUボードドライバクラス

Option Explicit

'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const ICUL1G_USE = 1                 '1CUL1Gボードの使用　  0：未使用、0以外：使用  <TesterがIP750EXならDefault:1にしておく>

#If ICUL1G_USE <> 0 Then

'############## DEFINE UNITS ##############
Private Const nA As Double = 0.000000001
Private Const uA As Double = 0.000001
Private Const mA As Double = 0.001
Private Const A As Double = 1#
Private Const nV As Double = 0.000000001
Private Const uV As Double = 0.000001
Private Const mV As Double = 0.001
Private Const V As Double = 1#
'##########################################

Private Const ALL_SITE As Long = -1
Private Const DEF_RANGE As Long = -1
Private Const BOARD_NAME As String = "dcICUL1G"
Private Const PPMU_VOLTAGE_RANGE As Double = 1 * V

Private Const FV_LIMIT_HIGH As Double = 6# * V
Private Const FV_LIMIT_LOW As Double = -1# * V
Private Const FI_LIMIT_HIGH As Double = 10# * mA
Private Const FI_LIMIT_LOW As Double = -10# * mA

Private Const CLAMP_V_HIGH As Double = 6.5 * V
Private Const CLAMP_V_LOW As Double = -1.5 * V

Private Const FV_RANGE_LIST As String = "None"
Private Const MV_RANGE_LIST As String = "None"
Private Const FI_RANGE_LIST As String = "20uA,200uA,2mA,10mA"
Private Const MI_RANGE_LIST As String = "20uA,200uA,2mA,10mA"

Private m_PinList As String
Private m_ResultsV  As Collection
Private m_ResultsI As Collection
Private m_ChannelList As CIcul1gChanList

Implements IVISVISource

'コンストラクタ
Private Sub Class_Initialize()
    Set m_ChannelList = New CIcul1gChanList
End Sub

'デストラクタ
Private Sub Class_Terminate()
    Set m_ChannelList = Nothing
    Set m_ResultsV = Nothing
    Set m_ResultsI = Nothing
End Sub

'I/F切り替え用の実装
Public Function AsIVISVISource() As IVISVISource
    Set AsIVISVISource = Me
End Function

'FVMIモード設定関数
Private Sub SetFVMI( _
    ByVal pPinList As String, _
    ByVal pForceV As Variant, _
    ByVal pMIRange As Icul1gPpmuCurrentRange, _
    Optional ByVal pSite As Long = ALL_SITE, _
    Optional ByVal pConnectOn As Boolean = True _
)

    'Error Check ************************************
    Const FunctionName = "SetFVMI_ICUL1G"
    If CheckIcul1gPpmuPinList(pPinList, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(pForceV, FV_LIMIT_LOW, FV_LIMIT_HIGH, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(pSite, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    '設定値のサイト配列作成
    Call ConvertVariableToArray(pForceV)
        
    'Site Loop **************************************
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = pSite Or pSite = ALL_SITE Then
            'Main Part ---------------------------------------
            TheHdw.ICUL1G.Pins(pPinList).PPMU.ForceVoltage(pMIRange) = pForceV(curSite)
            '-------------------------------------------------
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

    '接続設定
    If pConnectOn = True Then
        Call ConnectPins(pPinList, pSite)
    End If
    
End Sub

'FVMIモード設定関数(AllSite)
Private Sub SetFVMIMulti( _
    ByVal pPinList As String, _
    ByVal pForceV As Double, _
    ByVal pMIRange As Icul1gPpmuCurrentRange, _
    Optional ByVal pConnectOn As Boolean = True _
)

    'Error Check ************************************
    Const FunctionName = "SetFVMIMulti_ICUL1G"
    If CheckIcul1gPpmuPinList(pPinList, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(pForceV, FV_LIMIT_LOW, FV_LIMIT_HIGH, FunctionName) = False Then Stop: Exit Sub
    '************************************************
                
    'Main Part ---------------------------------------
    TheHdw.ICUL1G.Pins(pPinList).PPMU.ForceVoltage(pMIRange) = pForceV
    '-------------------------------------------------

    '接続設定(AllSite)
    If pConnectOn = True Then
        Call ConnectPinsMulti(pPinList)
    End If
    
End Sub

'FIMVモード設定関数
Private Sub SetFIMV( _
    ByVal pPinList As String, _
    ByVal pForceI As Variant, _
    Optional ByVal pClampVHi As Double = CLAMP_V_HIGH, _
    Optional ByVal pClampVLo As Double = CLAMP_V_LOW, _
    Optional ByVal pSite As Long = ALL_SITE, _
    Optional ByVal pFIRange As Icul1gPpmuCurrentRange = DEF_RANGE, _
    Optional ByVal pConnectOn As Boolean = True _
)
    
    'Error Check ************************************
    Const FunctionName = "SetFIMV_ICUL1G"
    If CheckIcul1gPpmuPinList(pPinList, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(pForceI, FI_LIMIT_LOW, FI_LIMIT_HIGH, FunctionName) = False Then Stop: Exit Sub
    If CheckClampVolage(pClampVLo, pClampVHi, FunctionName) = False Then Stop: Exit Sub
    If CheckMvModeCurrentRange(pFIRange, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(pSite, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    '設定値のサイト配列作成
    Call ConvertVariableToArray(pForceI)

    '電流レンジ値を求める
    Dim FIRangeBySite() As Icul1gPpmuCurrentRange
    ReDim FIRangeBySite(CountExistSite)
    Dim curSite As Long
    For curSite = 0 To CountExistSite
        If pFIRange = DEF_RANGE Then
            'Force値より最小レンジ値算出
            FIRangeBySite(curSite) = GetIcul1gPpmuIRange(pForceI(curSite))
        Else
            '指定レンジ値設定
            FIRangeBySite(curSite) = pFIRange
        End If
    Next curSite
    
    'Site Loop **************************************
    Dim siteStatus As LoopStatus
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = pSite Or pSite = ALL_SITE Then
            'Main Part ---------------------------------------
            With TheHdw.ICUL1G.Pins(pPinList).PPMU
                .ClampVHi = pClampVHi
                .ClampVLo = pClampVLo
                .ForceCurrent(FIRangeBySite(curSite)) = pForceI(curSite)
            End With
            '-------------------------------------------------
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************
    
    '接続設定
    If pConnectOn = True Then
        Call ConnectPins(pPinList, pSite)
    End If
        
End Sub

'FIMVモード設定関数(AllSite)
Private Sub SetFIMVMulti( _
    ByVal pPinList As String, _
    ByVal pForceI As Double, _
    Optional ByVal pClampVHi As Double = CLAMP_V_HIGH, _
    Optional ByVal pClampVLo As Double = CLAMP_V_LOW, _
    Optional ByVal pFIRange As Icul1gPpmuCurrentRange = DEF_RANGE, _
    Optional ByVal pConnectOn As Boolean = True _
)
    
    'Error Check ************************************
    Const FunctionName = "SetFIMVMulti_ICUL1G"
    If CheckIcul1gPpmuPinList(pPinList, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(pForceI, FI_LIMIT_LOW, FI_LIMIT_HIGH, FunctionName) = False Then Stop: Exit Sub
    If CheckClampVolage(pClampVLo, pClampVHi, FunctionName) = False Then Stop: Exit Sub
    If CheckMvModeCurrentRange(pFIRange, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    '電流レンジを求める
    Dim CurrentRange As Icul1gPpmuCurrentRange
    If pFIRange = DEF_RANGE Then
        'Force値より最小レンジ値算出
        CurrentRange = GetIcul1gPpmuIRange(pForceI)
    Else
        '指定レンジ値設定
        CurrentRange = pFIRange
    End If
            
    'Main Part ---------------------------------------
    With TheHdw.ICUL1G.Pins(pPinList).PPMU
        .ClampVHi = pClampVHi
        .ClampVLo = pClampVLo
        .ForceCurrent(CurrentRange) = pForceI
    End With
    '-------------------------------------------------
    
    '接続設定(AllSite)
    If pConnectOn = True Then
        Call ConnectPinsMulti(pPinList)
    End If
        
End Sub

'MVモード相当設定用関数（PPMUを0A印加状態に設定する）
'本当は本物のMVモード設定で実現したいが設定は、SetではなくMeasureにあるので他のリソースとリズムが合わない
Private Sub SetMV( _
    ByVal pPinList As String, _
    Optional ByVal pClampVHi As Double = CLAMP_V_HIGH, _
    Optional ByVal pClampVLo As Double = CLAMP_V_LOW, _
    Optional ByVal pSite As Long = ALL_SITE, _
    Optional pConnectOn As Boolean = True _
)
    
    Call SetFIMV(pPinList, 0 * A, pClampVHi, pClampVLo, pSite, icul1gPpmuCurrentRange20uA, pConnectOn)  '最小レンジ

End Sub

'MVモード相当設定用関数（PPMUを0A印加状態に設定する）(AllSite)
Private Sub SetMVMulti( _
    ByVal pPinList As String, _
    Optional ByVal pClampVHi As Double = CLAMP_V_HIGH, _
    Optional ByVal pClampVLo As Double = CLAMP_V_LOW, _
    Optional pConnectOn As Boolean = True _
)
    
    Call SetFIMVMulti(pPinList, 0 * A, pClampVHi, pClampVLo, icul1gPpmuCurrentRange20uA, pConnectOn)   '最小レンジ

End Sub

'PPMUを0V印加状態に設定する関数
Private Sub SetGND( _
    ByVal pPinList As String, _
    Optional ByVal pSite As Long = ALL_SITE, _
    Optional ByVal pConnectOn As Boolean = True _
)
    
    Call SetFVMI(pPinList, 0 * V, icul1gPpmuCurrentRange10mA, pSite, pConnectOn)  '最大レンジ
    
End Sub

'PPMUの電流測定レンジを変更する関数
Private Sub ChangeMIRange( _
ByVal pPinList As String, _
ByVal pMIRange As Icul1gPpmuCurrentRange, _
Optional ByVal pSite As Long = ALL_SITE _
)
    
    'Error Check ************************************
    Const FunctionName = "ChangeMIRange_ICUL1G"
    If CheckIcul1gPpmuPinList(pPinList, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(pSite, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    'Site Loop **************************************
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim ForceV As Double
    Dim ErrChans As String
    Dim ChanList As Collection
    Dim RetChanInfo As CIcul1gChanInfo
    
    ErrChans = Empty
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = pSite Or pSite = ALL_SITE Then
            'Main Part ---------------------------------------
            '指定サイトのチャンネルリストを取得
            Set ChanList = m_ChannelList.AtSiteAndPinList(pPinList, curSite)
            'チャンネルLOOP
            For Each RetChanInfo In ChanList
                With TheHdw.ICUL1G.Chans(RetChanInfo.ChanNumber, RetChanInfo.Icul1gChanType).PPMU
                    'FVモード判定
                    If .IsForcingVoltage = True Then
                        ForceV = .ForceVoltage(.CurrentRange)
                        .ForceVoltage(pMIRange) = ForceV
                    Else
                        'FIモードエラー
                        ErrChans = ErrChans & RetChanInfo.ChanNumber & ","
                    End If
                End With
            Next
            '-------------------------------------------------
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

    If ErrChans <> Empty Then
        Call OutputErrMsg("ICUL1G(" & ErrChans & ") is not MI Mode. (at" & FunctionName & ")")
    End If

End Sub

'PPMUの電流測定レンジを変更する関数(AllSite)
Private Sub ChangeMIRangeMulti( _
ByVal pPinList As String, _
ByVal pMIRange As Icul1gPpmuCurrentRange _
)
    
    'Error Check ************************************
    Const FunctionName = "ChangeMIRangeMulti_ICUL1G"
    If CheckIcul1gPpmuPinList(pPinList, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    'Activeサイトのチェック(Activeサイトが0の時には処理をしない)
    If TheExec.sites.ActiveCount = 0 Then
        Exit Sub
    End If
    
    '実行環境チェック
    Dim ForceVoltage As Double
    Dim ErrorMessage As String
    If CanChangeMIRangeMulti(pPinList, ForceVoltage, ErrorMessage) = False Then
        Call OutputErrMsg(ErrorMessage)
        Exit Sub
    End If
    
    '設定動作
    TheHdw.ICUL1G.Pins(pPinList).PPMU.ForceVoltage(pMIRange) = ForceVoltage
    
End Sub

'PPMUの電圧測定を行う関数
Private Sub MeasureV( _
    ByVal pPinName As String, _
    ByRef pRetResult() As Double, _
    ByVal pAvgNum As Long, _
    Optional ByVal pMvMode As Icul1gPpmuMVMode = icul1gPpmuMVModeNormal, _
    Optional ByVal pSite As Long = ALL_SITE _
)
            
    'Error Check ************************************
    Const FunctionName = "MeasureV_ICUL1G"
    If CheckIcul1gPpmuPinList(pPinName, FunctionName) = False Then Stop: Exit Sub
    If CheckIsSinglePin(pPinName, FunctionName) = False Then Stop: Exit Sub
    If CheckAvgNum(pAvgNum, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(pSite, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    '測定結果値の初期化処理
    Call InitializeResultArray(pRetResult, pSite)
     
    'Site Loop **************************************
    Dim curSite As Long
    Dim siteStatus As LoopStatus
    Dim ErrChans As String
    Dim ChanInfo As CIcul1gChanInfo
    
    '測定結果格納用の箱を用意
    Dim MeasureResult As CIcul1gPpmuResult
    Dim AllSiteResult() As Double
            
    ErrChans = Empty
    siteStatus = TheExec.sites.SelectFirst
    
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = pSite Or pSite = ALL_SITE Then
        '---- Main Part ----------------------------------
            'チャンネル情報の取得
            Set ChanInfo = m_ChannelList.AtSiteAndPin(pPinName, curSite)
            'FIMVモード設定確認
            If TheHdw.ICUL1G.Chans(ChanInfo.ChanNumber, ChanInfo.Icul1gChanType).PPMU.isForcingCurrent = False Then
                ErrChans = ErrChans & ChanInfo.ChanNumber & ","
            Else
                'アベレージ値設定
                TheHdw.ICUL1G.PPMU.Samples = pAvgNum
                'メーターリード
                Set MeasureResult = New CIcul1gPpmuResult
                Call TheHdw.ICUL1G.Pins(pPinName).PPMU.MeasureVoltages(MeasureResult.PinListData, pMvMode)
                'アベレージ値を1に設定
                TheHdw.ICUL1G.PPMU.Samples = 1
                '測定結果値の設定
'                Erase AllSiteResult
                AllSiteResult = MeasureResult.PinResult(pPinName)
                pRetResult(curSite) = AllSiteResult(curSite)
            End If
        '-------------------------------------------------
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

    If ErrChans <> Empty Then
        Call OutputErrMsg("ICUL1G(" & ErrChans & ") is not MV Mode (at" & FunctionName & ")")
    End If
    
End Sub

'PPMUの電流測定を行う関数
Private Sub MeasureI( _
    ByVal pPinName As String, _
    ByRef pRetResult() As Double, _
    ByVal pAvgNum As Long, _
    Optional ByVal pSite As Long = ALL_SITE _
)
            
    'Error Check ************************************
    Const FunctionName = "MeasureI_ICUL1G"
    If CheckIcul1gPpmuPinList(pPinName, FunctionName) = False Then Stop: Exit Sub
    If CheckIsSinglePin(pPinName, FunctionName) = False Then Stop: Exit Sub
    If CheckAvgNum(pAvgNum, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(pSite, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    '測定結果値の初期化処理
    Call InitializeResultArray(pRetResult, pSite)
         
    'Site Loop **************************************
    Dim curSite As Long
    Dim siteStatus As LoopStatus
    Dim ErrChans As String
    Dim ChanInfo As CIcul1gChanInfo
    
    '測定結果格納用の箱を用意
    Dim MeasureResult As CIcul1gPpmuResult
    Dim AllSiteResult() As Double
            
    ErrChans = Empty
    siteStatus = TheExec.sites.SelectFirst
    
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = pSite Or pSite = ALL_SITE Then
        '---- Main Part ----------------------------------
            'チャンネル情報の取得
            Set ChanInfo = m_ChannelList.AtSiteAndPin(pPinName, curSite)
            'FVMIモード設定確認
            If TheHdw.ICUL1G.Chans(ChanInfo.ChanNumber, ChanInfo.Icul1gChanType).PPMU.IsForcingVoltage = False Then
                ErrChans = ErrChans & ChanInfo.ChanNumber & ","
            Else
                'アベレージ値設定
                TheHdw.ICUL1G.PPMU.Samples = pAvgNum
                'メーターリード
                Set MeasureResult = New CIcul1gPpmuResult
                Call TheHdw.ICUL1G.Pins(pPinName).PPMU.MeasureCurrents(MeasureResult.PinListData)
                'アベレージ値を1に設定
                TheHdw.ICUL1G.PPMU.Samples = 1
                '測定結果値の設定
'                Erase AllSiteResult
                AllSiteResult = MeasureResult.PinResult(pPinName)
                pRetResult(curSite) = AllSiteResult(curSite)
            End If
        '-------------------------------------------------
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

    If ErrChans <> Empty Then
        Call OutputErrMsg("ICUL1G(" & ErrChans & ") is not MI Mode (at" & FunctionName & ")")
    End If
    
End Sub

'マルチピンの電圧測定用関数
Private Sub MeasureVMulti( _
ByVal pPinList As String, _
ByVal pAvgNum As Long, _
Optional ByVal pMvMode As Icul1gPpmuMVMode = icul1gPpmuMVModeNormal _
)
    
    'Error Check ************************************
    Const FunctionName = "MeasureVMulti_ICUL1G"
    If CheckIcul1gPpmuPinList(pPinList, FunctionName) = False Then Stop: Exit Sub
    If CheckAvgNum(pAvgNum, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    'Check MV Mode ************************************************
    Dim ChanList As Collection
    Set ChanList = m_ChannelList.AtPinList(pPinList)
    Dim Retinfo As CIcul1gChanInfo
    For Each Retinfo In ChanList
        If TheExec.sites.site(Retinfo.SiteNumber).Selected = True Then
            If TheHdw.ICUL1G.Chans(Retinfo.ChanNumber, Retinfo.Icul1gChanType).PPMU.IsForcingVoltage = True Then
                Call OutputErrMsg(pPinList & " Including not MV Mode Pins. (at" & FunctionName & ")")
                Exit Sub
            End If
        End If
    Next
    '**************************************************************
    
    'Measurement **************************************************
    Dim MeasureResult As CIcul1gPpmuResult
    Set MeasureResult = New CIcul1gPpmuResult
    TheHdw.ICUL1G.PPMU.Samples = pAvgNum
    With TheHdw.ICUL1G.Pins(pPinList).PPMU
        Call .MeasureVoltages(MeasureResult.PinListData, pMvMode)
    End With
    TheHdw.ICUL1G.PPMU.Samples = 1
    Set m_ResultsV = MeasureResult.PerPinResult
    '**************************************************************
    
End Sub

'マルチピンの電流測定用関数
Private Sub MeasureIMulti( _
ByVal pPinList As String, _
ByVal pAvgNum As Long _
)
    
    'Error Check ************************************
    Const FunctionName = "MeasureIMulti_ICUL1G"
    If CheckIcul1gPpmuPinList(pPinList, FunctionName) = False Then Stop: Exit Sub
    If CheckAvgNum(pAvgNum, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    'Check MI Mode ************************************************
    Dim ChanList As Collection
    Set ChanList = m_ChannelList.AtPinList(pPinList)
    Dim Retinfo As CIcul1gChanInfo
    For Each Retinfo In ChanList
        If TheExec.sites.site(Retinfo.SiteNumber).Selected = True Then
            If TheHdw.ICUL1G.Chans(Retinfo.ChanNumber, Retinfo.Icul1gChanType).PPMU.isForcingCurrent = True Then
                Call OutputErrMsg(pPinList & " Including not MI Mode Pins. (at" & FunctionName & ")")
                Exit Sub
            End If
        End If
    Next
    '**************************************************************
    
    'Measurement **************************************************
    Dim MeasureResult As CIcul1gPpmuResult
    Set MeasureResult = New CIcul1gPpmuResult
    TheHdw.ICUL1G.PPMU.Samples = pAvgNum
    With TheHdw.ICUL1G.Pins(pPinList).PPMU
        Call .MeasureCurrents(MeasureResult.PinListData)
    End With
    TheHdw.ICUL1G.PPMU.Samples = 1
    Set m_ResultsI = MeasureResult.PerPinResult
    '**************************************************************
    
End Sub

'MeasureVMultiで得た電流値を読み出す関数
Private Sub ReadMVMultiResult(ByVal pPinName As String, ByRef pResult() As Double)
    Dim status As Boolean
    status = ReadMultiResult(pPinName, pResult, m_ResultsV)
    If status = False Then
        Call OutputErrMsg("Not Found Result of " & pPinName & " at ReadMVMultiResult()")
        Stop
    End If
End Sub

'MeasureIMultiで得た電流値を読み出す関数
Private Sub ReadMIMultiResult(ByVal pPinName As String, ByRef pResult() As Double)
    Dim status As Boolean
    status = ReadMultiResult(pPinName, pResult, m_ResultsI)
    If status = False Then
        Call OutputErrMsg("Not Found Result of " & pPinName & " at ReadMIMultiResult()")
        Stop
    End If
End Sub

'PPMUを接続する関数
Private Sub ConnectPins( _
    ByVal pPinList As String, _
    Optional ByVal pSite As Long = ALL_SITE _
)
    
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    
    'Error Check ************************************
    Const FunctionName = "ConnectPins_ICUL1G"
    If CheckIcul1gPpmuPinList(pPinList, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(pSite, FunctionName) = False Then Stop: Exit Sub
    '************************************************

    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = pSite Or pSite = ALL_SITE Then
            'Main Part ---------------------------------------
            Call TheHdw.ICUL1G.Pins(pPinList).Connect
            '-------------------------------------------------
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

End Sub

'PPMUを接続する関数(AllSite)
Private Sub ConnectPinsMulti(ByVal pPinList As String)
    'Error Check ************************************
    Const FunctionName = "ConnectPinsMulti_ICUL1G"
    If CheckIcul1gPpmuPinList(pPinList, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    'Main Part ---------------------------------------
    Call TheHdw.ICUL1G.Pins(pPinList).Connect
    '-------------------------------------------------
End Sub

'PPMUを切り離す関数
Private Sub DisconnectPins( _
ByVal pPinList As String, _
Optional ByVal pSite As Long = ALL_SITE _
)
    
    'Error Check ************************************
    Const FunctionName = "DisconnectPins_ICUL1G"
    If CheckIcul1gPpmuPinList(pPinList, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(pSite, FunctionName) = False Then Stop: Exit Sub
    '************************************************

    'Site Loop **************************************
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = pSite Or pSite = ALL_SITE Then
            'Main Part ---------------------------------------
            'コンディションリセット処理
            With TheHdw.ICUL1G.Pins(pPinList)
                .Disconnect
                .PPMU.ForceCurrent(icul1gPpmuCurrentRange10mA) = 0 * A       '最大レンジ
                .PPMU.ForceVoltage(icul1gPpmuCurrentRangeAutoRange) = 0 * V
                .PPMU.ClampVHi = CLAMP_V_HIGH
                .PPMU.ClampVLo = CLAMP_V_LOW
            End With
            '-------------------------------------------------
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

End Sub

'PPMUを切り離す関数(AllSite)
Private Sub DisconnectPinsMulti(ByVal pPinList As String)
    'Error Check ************************************
    Const FunctionName = "DisconnectPinsMulti_ICUL1G"
    If CheckIcul1gPpmuPinList(pPinList, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    'Main Part ---------------------------------------
    'コンディションリセット処理
    With TheHdw.ICUL1G.Pins(pPinList)
        .Disconnect
        .PPMU.ForceCurrent(icul1gPpmuCurrentRange10mA) = 0 * A       '最大レンジ
        .PPMU.ForceVoltage(icul1gPpmuCurrentRangeAutoRange) = 0 * V
        .PPMU.ClampVHi = CLAMP_V_HIGH
        .PPMU.ClampVLo = CLAMP_V_LOW
    End With
    '-------------------------------------------------
End Sub

'ChangeMIRangeMultiを実行できる環境にあるかを確認する関数(True:OK)
Private Function CanChangeMIRangeMulti(ByVal pTargetPins As String, _
Optional ByRef pForceVoltage As Double, _
Optional ByRef pErrorMessage As String) As Boolean

    Const COMPARE_DIGIT As Long = 12 'DCシナリオのレンジバリデーションの精度と同等値

    Dim Snapshot As New CIcul1gSnapshotProvider
    Dim PpmuInfo As CIcul1gPpmuInfo

    '値初期化
    pForceVoltage = 0#
    pErrorMessage = Empty
    CanChangeMIRangeMulti = False
    
    For Each PpmuInfo In Snapshot.PpmuListAtPinList(pTargetPins)
        With PpmuInfo
            'Activeサイトのデータかを判定
            If IsActiveSite(.SiteNumber) = True Then
                'FVMIモード判定
                If .FVMode = False Then
                    'FailMode MIモードではない
                    CanChangeMIRangeMulti = False
                    pForceVoltage = 0#
                    pErrorMessage = "CH= " & .ChannelNumber & "(" & .StrChannelType & ")" & vbCrLf & "MV mode pin is present"
                    Exit Function
                End If
                'Force値初期値判定
                If CompareDblData(pForceVoltage, 0#, 0) = True Then
                    pForceVoltage = .ForceVoltage '基準値格納
                Else
                    'Force値比較処理
                    If CompareDblData(pForceVoltage, .ForceVoltage, COMPARE_DIGIT) = False Then
                        'FailMode Force値が同一でない
                        CanChangeMIRangeMulti = False
                        pForceVoltage = 0#
                        pErrorMessage = "CH= " & .ChannelNumber & "(" & .StrChannelType & ")" & vbCrLf & "Pin a different value of the ForceV exists"
                        Exit Function
                    End If
                End If
            End If
        End With
    Next
        
    If CompareDblData(pForceVoltage, 0#, 0) = True Then
        'FailMode アクティブサイトがひとつもない
        CanChangeMIRangeMulti = False
        pErrorMessage = "Active site does not exist"
        Exit Function
    Else
        'PassMode
        CanChangeMIRangeMulti = True
    End If

End Function

'電流レンジにAutoが設定されてないか確認する関数 (TRUE:OK)
'(AutoレンジはMIモードでのみ設定可能)
Private Function CheckMvModeCurrentRange(ByVal pCurrentRange As Icul1gPpmuCurrentRange, ByVal pFunctionName As String) As Boolean
    If pCurrentRange = icul1gPpmuCurrentRangeAutoRange Then
        CheckMvModeCurrentRange = False
        Call OutputErrMsg("AutoRange, FVMI mode can be set only " & pFunctionName & "().")
    Else
        CheckMvModeCurrentRange = True
    End If
End Function

'クランプ電圧の設定値を確認する関数(TRUE:OK)
Private Function CheckClampVolage(ByVal pLowValue As Double, ByVal pHighValue As Double, ByVal pFunctionName As String) As Boolean
    
    '上下関係のチェック
    If pLowValue >= pHighValue Then
        CheckClampVolage = False
        Call OutputErrMsg("Clamp voltage(" & "Hi= " & pHighValue & " ,Lo= " & pLowValue & ")" & vbCrLf _
        & "Voltage clamp, please set so that the Hi > Low  at " & pFunctionName & "().")
    Else
        CheckClampVolage = True
    End If
    
    '値の許容値チェック
    If (pLowValue < CLAMP_V_LOW) Or (pLowValue > CLAMP_V_HIGH) Or (pHighValue < CLAMP_V_LOW) Or (pHighValue > CLAMP_V_HIGH) Then
        CheckClampVolage = False
        Call OutputErrMsg("Clamp voltage(" & "Hi= " & pHighValue & " ,Lo= " & pLowValue & ")" & vbCrLf _
        & "must be between " & CLAMP_V_LOW & " and " & CLAMP_V_HIGH & " at " & pFunctionName & "().")
    Else
        CheckClampVolage = True
    End If

End Function

'ピンリストのピンがICUL1GのPPMUピンか否かを確認する関数 (TRUE:OK)
Private Function CheckIcul1gPpmuPinList(ByVal pPinList As String, ByVal pFunctionName As String) As Boolean
    CheckIcul1gPpmuPinList = TheHdw.ICUL1G.IsICUL1GPin(pPinList, icul1gPinPPMU)
    If CheckIcul1gPpmuPinList = False Then
        Call OutputErrMsg(pPinList & " is Invalid Channel Type at " & pFunctionName & "().")
    End If
End Function

'ピンリストに複数ピンが定義されているか否かを確認する関数(TRUE:OK)
Private Function CheckIsSinglePin(ByVal pPinList As String, ByVal pFunctionName As String) As Boolean
    Dim RetPins() As String
    Dim pinCount As Long
    Call TheExec.DataManager.DecomposePinList(pPinList, RetPins, pinCount)
    If pinCount <> 1 Then
        CheckIsSinglePin = False
        Call OutputErrMsg(pPinList & " is not single pin at " & pFunctionName & "().")
    Else
        CheckIsSinglePin = True
    End If
End Function

'Force電流の値から最小レンジ値を返す関数
Private Function GetIcul1gPpmuIRange(ByVal pCurrentValue As Double) As Icul1gPpmuCurrentRange
    Dim AbsForceI As Double
    AbsForceI = Abs(pCurrentValue)
    If AbsForceI <= 0.00002 Then          '20uA
        GetIcul1gPpmuIRange = icul1gPpmuCurrentRange20uA
    ElseIf AbsForceI <= 0.0002 Then       '200uA
        GetIcul1gPpmuIRange = icul1gPpmuCurrentRange200uA
    ElseIf AbsForceI <= 0.002 Then        '2mA
        GetIcul1gPpmuIRange = icul1gPpmuCurrentRange2mA
    Else                                  '10mA
        GetIcul1gPpmuIRange = icul1gPpmuCurrentRange10mA
    End If
End Function

'測定結果サイト配列の初期化処理関数
Private Sub InitializeResultArray(ByRef pResult() As Double, ByVal pSiteNumber As Long)
    Dim currentSite As Long
    If pSiteNumber = ALL_SITE Then
        For currentSite = 0 To CountExistSite Step 1
            pResult(currentSite) = 0
        Next currentSite
    Else
        If IsActiveSite(pSiteNumber) = False Then
            pResult(pSiteNumber) = 0
            Exit Sub
        End If
    End If
End Sub

'PPMUのForce値取得用関数
Private Function GetForceValue(ByRef pTarget As CIcul1gChanInfo) As Double
    'PPMUオブジェクトの取得
    Dim TargetChan As DriverICUL1G_Chan
    Set TargetChan = TheHdw.ICUL1G.Chans(pTarget.ChanNumber, pTarget.Icul1gChanType)
    Dim TargetPpmu As DriverICUL1G_PPMU_chan
    Set TargetPpmu = TargetChan.PPMU
    'PPMUモード別取得処理
    With TargetPpmu
        If .IsForcingVoltage = True Then
            'FVモード
            GetForceValue = .ForceVoltage(.CurrentRange)
        Else
            'FIモード
            GetForceValue = .ForceCurrent(.CurrentRange)
        End If
    End With
End Function

'PPMUのRange値取得用関数
Private Function GetRangeValue(ByRef pTarget As CIcul1gChanInfo) As Double
    'PPMUオブジェクトの取得
    Dim TargetChan As DriverICUL1G_Chan
    Set TargetChan = TheHdw.ICUL1G.Chans(pTarget.ChanNumber, pTarget.Icul1gChanType)
    Dim TargetPpmu As DriverICUL1G_PPMU_chan
    Set TargetPpmu = TargetChan.PPMU
    'PPMUモード別取得処理
    With TargetPpmu
        If .IsForcingVoltage = True Then
            'FVモード
            GetRangeValue = Icul1gPpmuCurrentRange2Value(.CurrentRange)
        Else
            'FIモード
            GetRangeValue = PPMU_VOLTAGE_RANGE
        End If
    End With
End Function

'PPMUのIRange値を値に変換する関数
Private Function Icul1gPpmuCurrentRange2Value(ByVal pIRange As Icul1gPpmuCurrentRange) As Double
    Select Case pIRange
        Case icul1gPpmuCurrentRangeSmartRange
            Icul1gPpmuCurrentRange2Value = DC_RANGE_SMART     'Smart
        Case icul1gPpmuCurrentRangeAutoRange
            Icul1gPpmuCurrentRange2Value = DC_RANGE_AUTO      'Auto
        Case icul1gPpmuCurrentRangeUnknown
            Call OutputErrMsg("IRange is Unknown Type at " & TypeName(Me) & "().")
        Case Else
            Icul1gPpmuCurrentRange2Value = TheHdw.ICUL1G.PPMU.CurrentRangeToValue(pIRange)
    End Select
End Function

'Measure*Mulitiで使用するコレクションを初期化する関数
Private Sub InitMultiDCResult()
    Set m_ResultsV = Nothing
    Set m_ResultsI = Nothing
End Sub


'######################## 以下 V/I Source I/Fの実装 #############################

Private Property Get IVISVISource_ForceVRangeList() As String
    IVISVISource_ForceVRangeList = FV_RANGE_LIST
End Property

Private Property Get IVISVISource_ForceIRangeList() As String
    IVISVISource_ForceIRangeList = FI_RANGE_LIST
End Property

Private Property Get IVISVISource_MeasVRangeList() As String
    IVISVISource_MeasVRangeList = MV_RANGE_LIST
End Property

Private Property Get IVISVISource_MeasIRangeList() As String
    IVISVISource_MeasIRangeList = MI_RANGE_LIST
End Property

Private Function IVISVISource_GetForceValue() As Collection
    Set IVISVISource_GetForceValue = Nothing
    'FAILサイトが存在するか確認し、存在する場合はエラーとする
    If CheckFailSiteExists("GetForceValue") = True Then
        Exit Function
    End If
    'ピンリスト作成
    Dim PinList() As String
    Dim pinCount As Long
    Call TheExec.DataManager.DecomposePinList(m_PinList, PinList, pinCount)
    'ピンLOOP
    Dim itemCount As Long
    Dim ChannelList As Collection
    Dim forceValue() As Double
    Set IVISVISource_GetForceValue = New Collection
    Dim Retinfo As CIcul1gChanInfo
    For itemCount = 0 To UBound(PinList) Step 1
        'チャンネルリスト取得
        Set ChannelList = m_ChannelList.AtPin(PinList(itemCount))
        'チャンネルリストLOOP
        ReDim forceValue(CountExistSite)
        For Each Retinfo In ChannelList
            forceValue(Retinfo.SiteNumber) = GetForceValue(Retinfo)
        Next
        'ピン名をキーに取得結果登録処理
        Call IVISVISource_GetForceValue.Add(forceValue, PinList(itemCount))
    Next itemCount
End Function

Private Function IVISVISource_GetForceVLimit() As Double()
    Dim tmpLimitVal(1) As Double
    tmpLimitVal(0) = FV_LIMIT_LOW
    tmpLimitVal(1) = FV_LIMIT_HIGH
    IVISVISource_GetForceVLimit = tmpLimitVal
End Function

Private Function IVISVISource_GetForceILimit() As Double()
    Dim tmpLimitVal(1) As Double
    tmpLimitVal(0) = FI_LIMIT_LOW
    tmpLimitVal(1) = FI_LIMIT_HIGH
    IVISVISource_GetForceILimit = tmpLimitVal
End Function

Private Function IVISVISource_GetMeasureRange() As Collection
    Set IVISVISource_GetMeasureRange = Nothing
    'FAILサイトが存在するか確認し、存在する場合はエラーとする
    If CheckFailSiteExists("GetMeasureRange") = True Then
        Exit Function
    End If
    'ピンリスト作成
    Dim PinList() As String
    Dim pinCount As Long
    Call TheExec.DataManager.DecomposePinList(m_PinList, PinList, pinCount)
    'ピンLOOP
    Dim itemCount As Long
    Dim ChannelList As Collection
    Dim RangeValue() As Double
    Set IVISVISource_GetMeasureRange = New Collection
    Dim Retinfo As CIcul1gChanInfo
    For itemCount = 0 To UBound(PinList) Step 1
        'チャンネルリスト取得
        Set ChannelList = m_ChannelList.AtPin(PinList(itemCount))
        'チャンネルリストLOOP
        ReDim RangeValue(CountExistSite)
        For Each Retinfo In ChannelList
            RangeValue(Retinfo.SiteNumber) = GetRangeValue(Retinfo)
        Next
        'ピン名をキーに取得結果登録処理
        Call IVISVISource_GetMeasureRange.Add(RangeValue, PinList(itemCount))
    Next itemCount
End Function

Private Property Get IVISVISource_BoardName() As String
    IVISVISource_BoardName = BOARD_NAME
End Property

Private Property Let IVISVISource_PinList(ByVal PinList As String)
    m_PinList = PinList
End Property

Private Property Get IVISVISource_PinList() As String
    IVISVISource_PinList = m_PinList
End Property

Private Sub IVISVISource_SetFVMI(ByVal ForceV As Variant, ByVal ClampI As Double, Optional ByVal site As Long = -1&, Optional ByVal ConnectOn As Boolean = True)
    Call SetFVMI(m_PinList, ForceV, GetIcul1gPpmuIRange(ClampI), site, ConnectOn)
End Sub

Private Sub IVISVISource_SetFVMIMulti(ByVal ForceV As Variant, ByVal ClampI As Double, Optional ByVal ConnectOn As Boolean = True)
   'site0ここで決めうちは危険 本来は、この関数にパラメータを渡す人がsite0の値を渡しているだけで
    'ライブラリは、単一の値のみの受付とし、それがどのサイトの値であるかに興味を持ってはならないはず
    Const USE_PARAMETER_SITE As Long = 0
    '使用する印加電圧の決定
    Dim ForceVoltage As Double
    If IsArray(ForceV) = True Then
        ForceVoltage = ForceV(USE_PARAMETER_SITE)
    Else
        ForceVoltage = CDbl(ForceV)
    End If
    Call SetFVMIMulti(m_PinList, ForceVoltage, GetIcul1gPpmuIRange(ClampI), ConnectOn)
End Sub

Private Sub IVISVISource_SetFIMV(ByVal ForceI As Variant, ByVal ClampV As Double, Optional ByVal site As Long = -1&, Optional ByVal ConnectOn As Boolean = True)
    'クランプ電圧は、下限は常にLOW(-1.5V)固定、上限はパラメータ値をそのまま採用する
    Call SetFIMV(m_PinList, ForceI, ClampV, CLAMP_V_LOW, site, , ConnectOn)
End Sub

Private Sub IVISVISource_SetFIMVMulti(ByVal ForceI As Variant, ByVal ClampV As Double, Optional ByVal ConnectOn As Boolean = True)
    'site0ここで決めうちは危険 本来は、この関数にパラメータを渡す人がsite0の値を渡しているだけで
    'ライブラリは、単一の値のみの受付とし、それがどのサイトの値であるかに興味を持ってはならないはず
    Const USE_PARAMETER_SITE As Long = 0
    '使用する印加電流の決定
    Dim ForceCurrent As Double
    If IsArray(ForceI) = True Then
        ForceCurrent = ForceI(USE_PARAMETER_SITE)
    Else
        ForceCurrent = CDbl(ForceI)
    End If
    'クランプ電圧は、下限は常にLOW(-1.5V)固定、上限はパラメータ値をそのまま採用する
    Call SetFIMVMulti(m_PinList, ForceCurrent, ClampV, CLAMP_V_LOW, , ConnectOn)
End Sub

Private Sub IVISVISource_SetMV(ByVal ClampV As Double, Optional ByVal site As Long = -1&, Optional ByVal ConnectOn As Boolean = True)
    Call SetMV(m_PinList, ClampV, CLAMP_V_LOW, site, ConnectOn)
End Sub

Private Sub IVISVISource_SetMVMulti(ByVal ClampV As Double, Optional ByVal ConnectOn As Boolean = True)
    Call SetMVMulti(m_PinList, ClampV, CLAMP_V_LOW, ConnectOn)
End Sub

Private Sub IVISVISource_SetGND(Optional ByVal site As Long = -1&, Optional ByVal ConnectOn As Boolean = True)
    Call SetGND(m_PinList, site, ConnectOn)
End Sub

Private Sub IVISVISource_ConnectPins(Optional ByVal site As Long = -1&)
    Call ConnectPins(m_PinList, site)
End Sub

Private Sub IVISVISource_DisconnectPins(Optional ByVal site As Long = -1&)
    Call DisconnectPins(m_PinList, site)
End Sub

Private Sub IVISVISource_DisconnectPinsMulti()
    Call DisconnectPinsMulti(m_PinList)
End Sub

Private Sub IVISVISource_ChangeMIRange(ByVal ClampI As Double, Optional ByVal site As Long = -1&)
    Call ChangeMIRange(m_PinList, GetIcul1gPpmuIRange(ClampI), site)
End Sub

Private Sub IVISVISource_ChangeMIRangeMulti(ByVal ClampI As Double)
    Call ChangeMIRangeMulti(m_PinList, GetIcul1gPpmuIRange(ClampI))
End Sub

Private Sub IVISVISource_MeasureV(retResult() As Double, ByVal avgNum As Long, Optional ByVal site As Long = -1&)
    Call MeasureV(m_PinList, retResult, avgNum, , site)
End Sub

Private Sub IVISVISource_MeasureI(retResult() As Double, ByVal avgNum As Long, Optional ByVal site As Long = -1&)
    Call MeasureI(m_PinList, retResult, avgNum, site)
End Sub

Private Sub IVISVISource_MeasureVMulti(ByVal avgNum As Long)
    Call InitMultiDCResult
    Call MeasureVMulti(m_PinList, avgNum)
End Sub

Private Sub IVISVISource_MeasureIMulti(ByVal avgNum As Long)
    Call InitMultiDCResult
    Call MeasureIMulti(m_PinList, avgNum)
End Sub

Private Sub IVISVISource_ReadMVMultiResult(retResult() As Double)
    Call ReadMVMultiResult(m_PinList, retResult)
End Sub

Private Sub IVISVISource_ReadMIMultiResult(retResult() As Double)
    Call ReadMIMultiResult(m_PinList, retResult)
End Sub

'################################################################################

#End If

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVISDps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   DPSボード制御用クラス
'
'目的:
'   Ⅰ:DPSボードの使い勝手向上
'   Ⅱ:V/ISourceとしてのI/F提供
'
'作成者:
'   SLSI今手
'
'XLibSTD_CommonDCModよりDPS関連機能を整理
'
'注意事項:
'GANG接続設定には対応していません。
'GANG接続用PinGPをPinListに設定するとエラーとなります
'
'
'Code Checked
'Comment Checked
'

Option Explicit

'############## DEFINE UNITS ##############
Private Const nA = 0.000000001
Private Const uA = 0.000001
Private Const mA = 0.001
Private Const A = 1
Private Const nV = 0.000000001
Private Const uV = 0.000001
Private Const mV = 0.001
Private Const V = 1
'##########################################

Private Const ALL_SITE = -1
Private Const BOARD_NAME = "dcDPS"

Private Const FV_LIMIT_HIGH = 10# * V
Private Const FV_LIMIT_LOW = 0# * V

Private Const FV_RANGE_LIST = "None" '0to10Vだが設定パラメータは無い
Private Const FI_RANGE_LIST = "None" 'FIMVモード設定は無い
Private Const MV_RANGE_LIST = "None" 'FIMVモード設定は無い
Private Const MI_RANGE_LIST = "50uA,500uA,10mA,100mA,1A"

Private m_PinList As String
Private m_ResultsI_DPS As Collection

Implements IVISVISource

'#Pass-OK
Public Function AsIVISVISource() As IVISVISource
'内容:
'   IVISVISource OBJ取得
'
'パラメータ:
'
'戻り値:
'   IVISVISource オブジェクト
'
'注意事項:
'
    Set AsIVISVISource = Me

End Function

'#Pass-Ok
Private Sub SetFVMI( _
    ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, _
    ByVal MIRange As DpsIRange, Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'   DPSを電圧印加状態に設定する｡
'
'パラメータ:
'    [PinList]    In   対象ピンリスト。
'    [ForceV]     In   印加電圧。配列指定可能。
'    [ClampI]     In   クランプ電流値。
'    [MIRange]    In   電流測定レンジ。
'    [Site]       In   サイト指定。オプション(Default -1)
'    [ConnectOn]  In   デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するDPSを電圧印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ForceVで印加電圧を指定。ForceVは数値 or サイト数分の配列。
'    ■数値の場合全サイト同じ値を設定｡配列の場合各サイトに対応する配列要素の値を設定｡
'    ■ClampIでクランプ電流を設定｡
'    ■クランプ電流は50mA～1A。
'    ■MIRangeで測定レンジを指定｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    
    
    'Error Check ************************************
    Const FunctionName = "SetFVMI_DPS"
    If CheckPinList(PinList, chDPS, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceV, 0 * V, 10 * V, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 50 * mA, 1 * A, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    Call ConvertVariableToArray(ForceV)
        
    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            With TheHdw.DPS.Pins(PinList)
                .CurrentRange = MIRange
                .CurrentLimit = ClampI
                .forceValue(dpsPrimaryVoltage) = ForceV(curSite)
            End With
            '-------------------------------------------------
            
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

    If ConnectOn = True Then
        Call ConnectPins(PinList, site)
    End If
    
End Sub
'#V21-Release
Private Sub SetFVMIMulti( _
    ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, _
    ByVal MIRange As DpsIRange, _
    Optional ByVal ConnectOn As Boolean = True _
)

'内容:
'   DPSを電圧印加状態に設定する｡(Site同時)
'
'パラメータ:
'    [PinList]    In   対象ピンリスト。
'    [ForceV]     In   印加電圧。配列指定可能。
'    [ClampI]     In   クランプ電流値。
'    [MIRange]    In   電流測定レンジ。
'    [ConnectOn]  In   デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するDPSを電圧印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ForceVで印加電圧を指定。ForceVは数値 or サイト数分の配列。
'    ■全サイト同じ値を設定｡
'    ■ClampIでクランプ電流を設定｡
'    ■クランプ電流は50mA～1A。
'    ■MIRangeで測定レンジを指定｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
'    Dim SiteStatus As loopStatus
    Dim curSite As Long
    Dim Channels() As Long
    
    
    'Error Check ************************************
    Const FunctionName = "SetFVMIMulti_DPS"
    If CheckPinList(PinList, chDPS, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceV, 0 * V, 10 * V, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 50 * mA, 1 * A, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
    'site0only設定************************************
    Call ConvertVariableToArray(ForceV)
    curSite = 0
    '************************************************
          
            'Main Part ---------------------------------------
 'Pins指定
            With TheHdw.DPS.Pins(PinList)
                .CurrentRange = MIRange
                .CurrentLimit = ClampI
                .forceValue(dpsPrimaryVoltage) = ForceV(curSite)
            End With
            
'Chans指定
'            Call GetActiveChanList(pinList, chDPS, Channels)
'            With TheHdw.BPMU.Chans(Channels)
'             .CurrentRange = MIRange
'                .CurrentLimit = ClampI
'                .forceValue(dpsPrimaryVoltage) = ForceV(CurSite)
'            End With
            '-------------------------------------------------
            

    If ConnectOn = True Then
        Call ConnectPinsMultiPins(PinList)
'        Call ConnectPinsMultiChans(pinList)
    End If
    
End Sub
'#Pass-Ok
Private Sub SetGND( _
    ByVal PinList As String, Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'   DPSを接地する｡
'
'パラメータ:
'    [PinList]     In   対象ピンリスト。
'    [Site]        In   サイト指定。オプション(Default -1)
'    [ConnectOn]   In   デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'詳細動作
'    ■PinListに対応するDPSを接地する｡ (OutputSourceをZeroVoltageに設定)
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
    Call SetFVMI(PinList, 0 * V, 1 * A, dps1a, site, ConnectOn)

End Sub

'#Pass-Ok
Private Sub DisconnectPins(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)
'内容:
'   DPSをデバイスから切り離す｡
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'    [Site]      In   サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するDPSをデバイスから切り離す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■切り離した後、0V印加状態に設定する(接続はしない)。
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    
    'Error Check ************************************
    Const FunctionName = "DisconnectPins_DPS"
    If CheckPinList(PinList, chDPS, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            Call GetChanList(PinList, curSite, chDPS, Channels)
            With TheHdw.DPS.Chans(Channels)
                .ForceRelayClosed = False
                .SenseRelayClosed = False
            End With
            
            With TheHdw.DPS.Pins(PinList)
                .forceValue(dpsPrimaryVoltage) = 0 * V
            End With
            '-------------------------------------------------
            
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

    Call DisconnectGangedPins(PinList, site)
    
End Sub
'#V21-Release
Private Sub DisconnectPinsMulti(ByVal PinList As String)
'内容:
'   DPSをデバイスから切り離す｡(Site同時)
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するDPSをデバイスから切り離す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■切り離した後、0V印加状態に設定する(接続はしない)。
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long
    Dim Channels() As Long
    
    'Error Check ************************************
    Const FunctionName = "DisconnectPinsMulti_DPS"
    If CheckPinList(PinList, chDPS, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    

        
    'Main Part ---------------------------------------
    Call GetActiveChanList(PinList, chDPS, Channels)
    With TheHdw.DPS.Chans(Channels)
    .ForceRelayClosed = False
    .SenseRelayClosed = False
    End With
            
    With TheHdw.DPS.Pins(PinList)
    .forceValue(dpsPrimaryVoltage) = 0 * V
    End With
    '-------------------------------------------------
 

    Call DisconnectGangedPins(PinList, ALL_SITE)
    
End Sub
'#Pass-Ok
Private Sub ConnectPins(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)
'内容:
'   DPSをデバイスに接続する｡
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'    [Site]      In   サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するDPSをデバイスに接続する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    
    'Error Check ************************************
    Const FunctionName = "ConnectPins_DPS"
    If CheckPinList(PinList, chDPS, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            Call GetChanList(PinList, curSite, chDPS, Channels)
            With TheHdw.DPS.Chans(Channels)
                .ForceRelayClosed = True
                .SenseRelayClosed = True
            End With
            '-------------------------------------------------
            
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

End Sub
'#V21-Release
Private Sub ConnectPinsMultiChans(ByVal PinList As String)
'内容:
'   DPSをデバイスに接続する｡(サイト同時)
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するDPSをデバイスに接続する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long
    Dim Channels() As Long
    
    'Error Check ************************************
    Const FunctionName = "ConnectPinsMulti_DPS"
    If CheckPinList(PinList, chDPS, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
            
            'Main Part ---------------------------------------
            Call GetActiveChanList(PinList, chDPS, Channels)
            With TheHdw.DPS.Chans(Channels)
                .ForceRelayClosed = True
                .SenseRelayClosed = True
            End With
            '-------------------------------------------------
            

End Sub
'#V21-Release
Private Sub ConnectPinsMultiPins(ByVal PinList As String)
'内容:
'   DPSをデバイスに接続する｡(サイト同時)
'   中身はchans
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するDPSをデバイスに接続する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long
    Dim Channels() As Long
    
    'Error Check ************************************
    Const FunctionName = "ConnectPinsMulti_DPS"
    If CheckPinList(PinList, chDPS, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
            
            'Main Part ---------------------------------------
            Call GetActiveChanList(PinList, chDPS, Channels)
            With TheHdw.DPS.Chans(Channels)
                .ForceRelayClosed = True
                .SenseRelayClosed = True
            End With
            '-------------------------------------------------
            

End Sub
'#Pass-Ok
Private Sub ChangeMIRange( _
    ByVal PinList As String, ByVal ClampI As Double, _
    ByVal MIRange As DpsIRange, Optional ByVal site As Long = ALL_SITE _
)
'内容:
'   電圧印加状態のDPSの電流測定レンジを変更する｡
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'    [ClampI]    In   クランプ電流値。
'    [MIRange]   In   電流測定レンジ。
'    [Site]      In   サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するDPSの印加電圧を変えず､電流測定レンジを変更する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ClampIでクランプ電流を設定｡
'    ■クランプ電流は50mA～1A。
'    ■MIRangeで測定レンジを指定｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long
    
    'Error Check ************************************
    Const FunctionName = "ChangeMIRange_DPS"
    If CheckPinList(PinList, chDPS, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 50 * mA, 1 * A, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            Call GetChanList(PinList, curSite, chDPS, Channels)
            For chan = 0 To UBound(Channels)
                With TheHdw.DPS.Chans(Channels(chan))
                    .CurrentRange = MIRange
                    .CurrentLimit = ClampI
                End With
            Next chan
            '-------------------------------------------------
            
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

End Sub
'#V21-Release
Private Sub ChangeMIRangeMulti( _
    ByVal PinList As String, ByVal ClampI As Double, _
    ByVal MIRange As DpsIRange _
)

'内容:
'   電圧印加状態のDPSの電流測定レンジを変更する｡(Site同時)
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'    [ClampI]    In   クランプ電流値。
'    [MIRange]   In   電流測定レンジ。

'
'戻り値:
'
'注意事項:
'    ■PinListに対応するDPSの印加電圧を変えず､電流測定レンジを変更する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ClampIでクランプ電流を設定｡
'    ■クランプ電流は50mA～1A。
'    ■MIRangeで測定レンジを指定｡
'    ■非アクティブサイトに対しては何もしない｡
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long
    Dim Channels() As Long
'    Dim chan As Long
    
    'Error Check ************************************
    Const FunctionName = "ChangeMIRangeMulti_DPS"
    If CheckPinList(PinList, chDPS, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 50 * mA, 1 * A, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    

        
            'Main Part ---------------------------------------
'Pins指定
                With TheHdw.DPS.Pins(PinList)
                    .CurrentRange = MIRange
                    .CurrentLimit = ClampI
                End With
                       
'Chans指定
'            Call GetActiveChanList(pinList, chDPS, Channels)
'                With TheHdw.DPS.Chans(Channels)
'                    .CurrentRange = MIRange
'                    .CurrentLimit = ClampI
'                End With

            '-------------------------------------------------



End Sub
'#Pass-Ok
Private Sub MeasureI(ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, Optional ByVal site As Long = ALL_SITE)
'内容:
'    DPSの電流値を測定する｡
'
'パラメータ:
'    [PinName]      In   対象ピン名。
'    [retResult()]  Out  結果を格納する配列。(サイト数分の要素が必要)
'    [AvgNum]       In   平均回数。
'    [Site]         In   サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinNameに対応するAPMUの電流値を測定する｡
'    ■PinNameは１ピンのみ。複数ピン、ピングループ(ギャングピンを除く)の指定は不可。
'    ■retResultに結果を返す｡
'    ■AvgNumの数だけサンプリングして､その平均を返す｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■Siteが-1以外の場合、対象サイト以外のretResult()の値は変化させない。
'    ■非アクティブサイトの結果は0を返す。Siteが-1の時も同様。
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim Samples() As Double
    Dim n As Long
        
    'Error Check ************************************
    Const FunctionName = "MeasureI_DPS"
    If CheckPinList(PinName, chDPS, FunctionName) = False Then Stop: Exit Sub
    If CheckSinglePins(PinName, chDPS, FunctionName) = False Then Stop: Exit Sub
    If CheckResultArray(retResult, FunctionName) = False Then Stop: Exit Sub
    If CheckAvgNum(avgNum, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    If site = ALL_SITE Then
        For curSite = 0 To CountExistSite
            retResult(curSite) = 0
        Next curSite
    Else
        If IsActiveSite(site) = False Then
            retResult(site) = 0
            Exit Sub
        End If
    End If
    
    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
            
            'Main Part ---------------------------------------
            Call GetChanList(PinName, curSite, chDPS, Channels)
            
            With TheHdw.DPS.Pins(PinName)
                TheHdw.DPS.Samples = avgNum
                Call .MeasureCurrents(.CurrentRange, Samples)
                TheHdw.DPS.Samples = 1
                retResult(curSite) = 0
                For n = 0 To UBound(Samples)
                    retResult(curSite) = retResult(curSite) + Samples(n)
                Next n
                retResult(curSite) = retResult(curSite) / avgNum
            End With
            '-------------------------------------------------
        
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

End Sub

'#Pass-Ok
Private Sub MeasureIMulti(ByVal PinList As String, ByVal avgNum As Long)
'内容:
'   DPSの電流値を測定する｡
'
'パラメータ:
'    [PinList]   In   対象ピン名。
'    [AvgNum]    In   平均回数。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するピンの電流値を測定する｡結果はReadMIMultiResultを使って読み出す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ギャングピンは対応しない｡
'    ■AvgNumの数だけサンプリングして､その平均を返す｡
'    ■全サイトが対象になる｡非アクティブサイトの結果は0になる｡
'
    Dim Channels() As Long
    Dim Samples() As Double
    Dim pinNames() As String
    Dim PerPinResults() As Variant
    Dim curSite As Long
    Dim Pin As Long
    Dim n As Long
    Dim i As Long
        
    'Error Check ************************************
    Const FunctionName = "MeasureIMulti_DPS"
    If CheckPinList(PinList, chDPS, FunctionName) = False Then Stop: Exit Sub
    If CheckAvgNum(avgNum, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    'Check Ganged Pins ********************************************
    Call GetActiveChanList(PinList, chDPS, Channels)
    Call SeparatePinList(PinList, pinNames)
    If (UBound(pinNames) + 1) * CountActiveSite <> UBound(Channels) + 1 Then
        Call OutputErrMsg(PinList & " Including Ganged Pins. (at MeasureIMulti_DPS)")
        Exit Sub
    End If
    '**************************************************************
        
    'Measurement **************************************************
    With TheHdw.DPS.Pins(PinList)
        TheHdw.DPS.Samples = avgNum
        Call .MeasureCurrents(.CurrentRange, Samples)
        TheHdw.DPS.Samples = 1
    End With
    '**************************************************************
    
    'Regist Results ***********************************************
    'Initialize ----------------------------------------------
    PerPinResults = CreateEmpty2DArray(UBound(pinNames), CountExistSite)
    '---------------------------------------------------------
    
    'Summation -----------------------------------------------
    i = 0
    For Pin = 0 To UBound(pinNames)
        For curSite = 0 To CountExistSite
            If IsActiveSite(curSite) Then
                For n = 1 To avgNum
                    PerPinResults(Pin)(curSite) = PerPinResults(Pin)(curSite) + Samples(i)
                    i = i + 1
                Next n
           End If
        Next curSite
    Next Pin
    
    '---------------------------------------------------------
    
    'Average & Regist ----------------------------------------
    Set m_ResultsI_DPS = New Collection
    For Pin = 0 To UBound(pinNames)
        For curSite = 0 To CountExistSite
            PerPinResults(Pin)(curSite) = PerPinResults(Pin)(curSite) / avgNum
        Next curSite
        Call m_ResultsI_DPS.Add(PerPinResults(Pin), pinNames(Pin))
    Next Pin
    '---------------------------------------------------------
    '**************************************************************

End Sub

'#Pass-Ok
Private Sub ReadMIMultiResult(ByVal PinName As String, ByRef retResult() As Double)
'内容:
'   MeasureIMultiで得た電流値を読み出す｡
'
'パラメータ:
'    [PinName]       In   対象ピン名。
'    [retResult()]   Out  結果を格納する配列。(サイト数分の要素が必要)
'
'戻り値:
'
'注意事項:
'    ■PinNameは最後に使ったMeasureIMultiのピンリストに含まれるピン名を指定｡
'    ■retResultに結果を返す｡
'
    Dim status As Boolean
    
    status = ReadMultiResult(PinName, retResult, m_ResultsI_DPS)

    If status = False Then
        Call OutputErrMsg("Not Found Result of " & PinName & " at ReadMIMultiResult()")
        Stop
    End If
    
End Sub

'#Pass-Ok
Private Sub DisconnectGangedPins(ByVal PinList As String, ByVal site As Long)

    Dim pinArr() As String
    Dim pinNum As Long
    Dim Pin As Long
    Dim TmpPinList As String
    Dim Channels() As Long
    Dim curSite As Long
    Dim chan As Long
    
    If site <> ALL_SITE Then
        If TheExec.sites.site(site).Active = False Then Exit Sub
    End If
    
    Call TheExec.DataManager.DecomposePinList(PinList, pinArr, pinNum)
    
    TmpPinList = pinArr(0)
    For Pin = 1 To UBound(pinArr)
        TmpPinList = TmpPinList & "," & pinArr(Pin)
    Next Pin
    
    For curSite = 0 To CountExistSite
        If curSite = site Or (site = ALL_SITE And IsActiveSite(curSite) = True) Then
            Call GetChanList(TmpPinList, curSite, chDPS, Channels)
            For chan = 0 To UBound(Channels) - 1
                TheHdw.DPS.GangedChannels(Channels(chan), Channels(chan + 1)) = False
            Next
        End If
    Next curSite
    
End Sub

'#Pass-Ok
Private Sub InitMultiDCResult()
'内容:
'   Measure＊Mulitiで使用するコレクションを初期化
'
    Set m_ResultsI_DPS = Nothing

End Sub

'#Pass-Ok
Private Function GetDpsMIRange(ByVal ClampI As Double) As DpsIRange

    Dim AbsClampI As Double

    AbsClampI = Abs(ClampI)

    If AbsClampI <= 0.00005 Then        '50uA
        GetDpsMIRange = dps50uA
    ElseIf AbsClampI <= 0.0005 Then     '500uA
        GetDpsMIRange = dps500ua
    ElseIf AbsClampI <= 0.01 Then       '10mA
        GetDpsMIRange = dps10ma
    ElseIf AbsClampI <= 0.1 Then        '100mA
        GetDpsMIRange = dps100mA
    Else
        GetDpsMIRange = dps1a
    End If

End Function

'#Pass-Ok
Private Function GetDpsMeasureRange() As Collection
'DPSの測定レンジを取得する
'DPSはMIモードのみ

    Dim pinNameArr() As String
    Dim chListArr() As Long
    Dim pinLpCnt As Long
    Dim siteLpCnt As Long
    Dim tmpMeasureRng() As Double
    Dim tmpIrange As DpsIRange
    Dim tmpRngVal As Collection
                
    Const FUNCTION_NAME = "GetDpsMeasureRange"
                                                     
    '取得レンジ格納用変数
    ReDim tmpMeasureRng(CountExistSite)
    
    '最終取得結果用コレクションを準備
    Set tmpRngVal = New Collection
                
    'カンマ区切りのピンリストを配列形式に変換
    Call SeparatePinList(m_PinList, pinNameArr)
    
    'ピンLOOP
    For pinLpCnt = 0 To UBound(pinNameArr)
        'Pinリストからマルチサイト分のチャンネルリストを作成
        Call GetChanList(pinNameArr(pinLpCnt), ALL_SITE, GetChanType(pinNameArr(pinLpCnt)), chListArr)
                                
        'サイトループ(CHループ)
        For siteLpCnt = 0 To CountExistSite
            '現在の電流レンジ値をTERADYNE-APIより取得
            tmpIrange = TheHdw.DPS.Chans(chListArr(siteLpCnt)).CurrentRange
            tmpMeasureRng(siteLpCnt) = DpsIRangeToValue(tmpIrange)
        Next siteLpCnt
        
        'ピン名をキーにレンジ値の取得結果をコレクションに登録
        Call tmpRngVal.Add(tmpMeasureRng, pinNameArr(pinLpCnt))
    
    Next pinLpCnt
     
#If GANG_DPS_USE = 999 Then
'---------------------------------------------------------------------------
'取得結果算出処理
    
    'ピンリストの中身が、GangのGPが含まれているとき（並列接続数を算出し、測定レンジ値を計算する）
    If IsGangPinlist(m_PinList, GetChanType(m_PinList)) = True Then
                
        'ギャング用のGPがPinListに設定されているときは単一GPかどうかを確認
        Dim gangPinListArr() As String
        Call ConvertStrPinListToArrayPinList(m_PinList, gangPinListArr) 'ピングループを展開せずにカンマ区切り形式の配列に変換
        
        'Gangのマルチピンには対応できていないのでエラーにする
        If UBound(gangPinListArr) <> 0 Then
            Call MsgBox("Ganged pin group is included in the pinlist.(" & m_PinList & ")  at " & FUNCTION_NAME & "()", vbCritical, FUNCTION_NAME)
            Set GetDpsMeasureRange = Nothing
            Exit Function
        End If
                                                                             
        '★ギャング対象GPのメンバーはすべて同じ状態と信じて
        '  先頭PINのレンジ値と印加モードのみを調べる
        ' （本来はすべてのメンバーの状態から判断を行うのが正しい）
        Dim tmpGangMeasureRng() As Double
        tmpGangMeasureRng = tmpRngVal.Item(pinNameArr(0))

        '★モード別処理（MIモード時のみ電流レンジ×並列接続数を演算）
        '  (ここでもサイト0の値のみを確認してモードを判断するが
        '  本来はすべてのサイトのモード値を確認するのが正しい
        '  上記のメンバー確認のタイミングですべて確認済みのものを
        '  判定するのが良さそう。)
        Dim tmpGangRngVal As Collection
        Set tmpGangRngVal = New Collection

        'サイトループ(CHループ)で測定レンジ値×並列接続数を計算（MIモード）
        For siteLpCnt = 0 To CountExistSite
            tmpGangMeasureRng(siteLpCnt) = tmpGangMeasureRng(siteLpCnt) * (UBound(pinNameArr) + 1)
        Next siteLpCnt
        'Gang用のGp名をキーに測定レンジ結果を格納
        Call tmpGangRngVal.Add(tmpGangMeasureRng, m_PinList)
                
        'Gang用に計算された測定レンジ値を取得結果として返す
        Set GetDpsMeasureRange = tmpGangRngVal
                
    'ピンリストの中身が通常ピンのとき(取得結果をそのまま返す)
    Else
        '取得結果を関数の戻り値としてコレクションで返す
        Set GetDpsMeasureRange = tmpRngVal

    End If
#Else
    '取得結果を関数の戻り値としてコレクションで返す
    Set GetDpsMeasureRange = tmpRngVal
#End If
                
End Function

'#Pass-Ok
Private Function GetDpsForceValue() As Collection
'DPSのForce値を取得する
'DPSはFVモードのみ

    Dim pinNameArr() As String
    Dim chListArr() As Long
    Dim pinLpCnt As Long
    Dim siteLpCnt As Long
    Dim tmpForceValue() As Double
    Dim tmpIrange As DpsIRange
    Dim tmpForceVal As Collection
    Dim retForceVal As Variant
                
    Const FUNCTION_NAME = "GetDpsForceValue"
                                     
    '取得レンジ格納用変数
    ReDim tmpForceValue(CountExistSite)
    
    '最終取得結果用コレクションを準備
    Set tmpForceVal = New Collection
                
    'カンマ区切りのピンリストを配列形式に変換
    Call SeparatePinList(m_PinList, pinNameArr)
    
    'ピンLOOP
    For pinLpCnt = 0 To UBound(pinNameArr)
        
        'Pinリストからマルチサイト分のチャンネルリストを作成
        Call GetChanList(pinNameArr(pinLpCnt), ALL_SITE, GetChanType(pinNameArr(pinLpCnt)), chListArr)
        'サイトループ(CHループ)
        For siteLpCnt = 0 To CountExistSite
            
            With TheHdw.DPS.Chans(chListArr(siteLpCnt))
'本来は設定されているモードに応じたForce値を取得するのが正しいが、
'OutputSourceのプロパティ参照結果が正しく得られないので、
'ISのテストで常に使用している[dpsPrimaryVoltage]の値を決めうちで取得する。
'
            '----本来行うべき処理-----------------------------------------
'                'Forceモードを取得しモード別のForce値を取得する(本来行うべき処理)
'                If .OutputSource <> dpsZeroVoltage Then
'                    retForceVal = .forceValue(.OutputSource)
'                    tmpForceValue(siteLpCnt) = retForceVal(0)
'                Else
'                    tmpForceValue(siteLpCnt) = 0#   'ZeroVoltage MODEのとき
'                End If
            '----本来行うべき処理-----------------------------------------
            
            '----暫定処理(dpsPrimaryVoltageの値をForceの値とする)
                retForceVal = .forceValue(dpsPrimaryVoltage)
                tmpForceValue(siteLpCnt) = retForceVal(0)
            '----暫定処理----------------------------------------
            End With
        
        Next siteLpCnt
        
        'ピン名をキーにレンジ値の取得結果をコレクションに登録
        Call tmpForceVal.Add(tmpForceValue, pinNameArr(pinLpCnt))
    
    Next pinLpCnt
       
#If GANG_DPS_USE = 999 Then
'---------------------------------------------------------------------------
'取得結果算出処理
    
    'ピンリストの中身が、GangのGPが含まれているとき（並列接続数を算出し、測定レンジ値を計算する）
    If IsGangPinlist(m_PinList, GetChanType(m_PinList)) = True Then
                
        'ギャング用のGPがPinListに設定されているときは単一GPかどうかを確認
        Dim gangPinListArr() As String
        Call ConvertStrPinListToArrayPinList(m_PinList, gangPinListArr) 'ピングループを展開せずにカンマ区切り形式の配列に変換
        
        'Gangのマルチピンには対応できていないのでエラーにする
        If UBound(gangPinListArr) <> 0 Then
            Call MsgBox("Ganged pin group is included in the pinlist.(" & m_PinList & ")  at " & FUNCTION_NAME & "()", vbCritical, FUNCTION_NAME)
            Set GetDpsForceValue = Nothing
            Exit Function
        End If
                                                                             
        Dim tmpGangForceVal As Collection
        Set tmpGangForceVal = New Collection
        'Gang用のGp名をキーにForce値を格納
        Call tmpGangForceVal.Add(tmpForceVal.Item(pinNameArr(0)), m_PinList)
        Set GetDpsForceValue = tmpGangForceVal
                
    'ピンリストの中身が通常ピンのとき(取得結果をそのまま返す)
    Else
        '取得結果を関数の戻り値としてコレクションで返す
        Set GetDpsForceValue = tmpForceVal
    End If
#Else
    '取得結果を関数の戻り値としてコレクションで返す
    Set GetDpsForceValue = tmpForceVal
#End If
                
End Function

'#Pass-OK
Private Function DpsIRangeToValue(ByVal Irange As DpsIRange) As Double
'DPSのIRangeをDouble値に変換

    Select Case Irange
        Case dpsAutoRange
            DpsIRangeToValue = DC_RANGE_AUTO  '"AutoRange"
        
        Case dpsSmartRange
            DpsIRangeToValue = DC_RANGE_SMART '"SmartRange"
        
        Case Else
            DpsIRangeToValue = TheHdw.DPS.IRangeToValue(Irange)
    End Select

End Function

'######################## 以下 V/I Source I/Fの実装 #############################

'#Pass-Ok
Private Property Get IVISVISource_ForceVRangeList() As String
    IVISVISource_ForceVRangeList = FV_RANGE_LIST
End Property

'#Pass-Ok
Private Property Get IVISVISource_ForceIRangeList() As String
    IVISVISource_ForceIRangeList = FI_RANGE_LIST
End Property

'#Pass-Ok
Private Property Get IVISVISource_MeasVRangeList() As String
    IVISVISource_MeasVRangeList = MV_RANGE_LIST
End Property

'#Pass-Ok
Private Property Get IVISVISource_MeasIRangeList() As String
    IVISVISource_MeasIRangeList = MI_RANGE_LIST
End Property

'#Pass-Ok
Private Function IVISVISource_GetForceValue() As Collection

    Set IVISVISource_GetForceValue = Nothing
    
    'FAILサイトが存在するか確認し、存在する場合はエラーとする
    If CheckFailSiteExists("GetForceValue") = True Then
        Exit Function
    End If

    Set IVISVISource_GetForceValue = GetDpsForceValue

End Function

'#Pass-Ok
Private Function IVISVISource_GetForceVLimit() As Double()

    Dim tmpLimitVal(1) As Double

    tmpLimitVal(0) = FV_LIMIT_LOW
    tmpLimitVal(1) = FV_LIMIT_HIGH

    IVISVISource_GetForceVLimit = tmpLimitVal

End Function

'#Pass-Ok
Private Function IVISVISource_GetForceILimit() As Double()
    
     Call OutputErrMsg("DPS is not support FI mode")
        
End Function

'#Pass-OK
Private Function IVISVISource_GetMeasureRange() As Collection
    
    Set IVISVISource_GetMeasureRange = Nothing
    
    'FAILサイトが存在するか確認し、存在する場合はエラーとする
    If CheckFailSiteExists("GetMeasureRange") = True Then
        Exit Function
    End If
    
    Set IVISVISource_GetMeasureRange = GetDpsMeasureRange

End Function

'#Pass-Ok
Private Property Get IVISVISource_BoardName() As String
    
    IVISVISource_BoardName = BOARD_NAME

End Property

'#Pass-Ok
Private Property Get IVISVISource_PinList() As String
    
    IVISVISource_PinList = m_PinList

End Property

'#Pass-Ok
Private Property Let IVISVISource_PinList(ByVal PinList As String)
    
#If GANG_DPS_USE <> 999 Then
    'ギャングピンの設定はエラーとする
    If IsGangPinlist(PinList, GetChanType(PinList)) = True Then
        Call OutputErrMsg(PinList & " Including Ganged Pins. (at DPS_PinList)")
        Exit Property
    End If
#End If
    
    m_PinList = PinList

End Property

'#Pass-Ok
Private Sub IVISVISource_SetFVMI(ByVal ForceV As Variant, ByVal ClampI As Double, Optional ByVal site As Long = -1&, Optional ByVal ConnectOn As Boolean = True)
        
    Call SetFVMI(m_PinList, ForceV, CreateLimit(Abs(ClampI), 50 * mA, 1 * A), GetDpsMIRange(ClampI), site, ConnectOn)

End Sub
'#V21-Release
Private Sub IVISVISource_SetFVMIMulti(ByVal ForceV As Variant, ByVal ClampI As Double, Optional ByVal ConnectOn As Boolean = True)
        
    Call SetFVMIMulti(m_PinList, ForceV, CreateLimit(Abs(ClampI), 50 * mA, 1 * A), GetDpsMIRange(ClampI), ConnectOn)

End Sub
'#Pass-Ok
Private Sub IVISVISource_SetFIMV(ByVal ForceI As Variant, ByVal ClampV As Double, Optional ByVal site As Long = -1&, Optional ByVal ConnectOn As Boolean = True)
    
    Call OutputErrMsg(m_PinList & " is Invalid Pin List at SetFIMV()")
    Stop

End Sub
'#V21-Release
Private Sub IVISVISource_SetFIMVMulti(ByVal ForceI As Variant, ByVal ClampV As Double, Optional ByVal ConnectOn As Boolean = True)
    
    Call OutputErrMsg(m_PinList & " is Invalid Pin List at SetFIMVMulti()")
    Stop

End Sub
'#Pass-Ok
Private Sub IVISVISource_SetMV(ByVal ClampV As Double, Optional ByVal site As Long = -1&, Optional ByVal ConnectOn As Boolean = True)

    Call OutputErrMsg(m_PinList & " is Invalid Pin List at SetMV()")
    Stop

End Sub
'#V21-Release
Private Sub IVISVISource_SetMVMulti(ByVal ClampV As Double, Optional ByVal ConnectOn As Boolean = True)

    Call OutputErrMsg(m_PinList & " is Invalid Pin List at SetMVMulti()")
    Stop

End Sub
'#Pass-Ok
Private Sub IVISVISource_SetGND(Optional ByVal site As Long = -1&, Optional ByVal ConnectOn As Boolean = True)
    
    Call SetGND(m_PinList, site, ConnectOn)
    
End Sub

'#Pass-Ok
Private Sub IVISVISource_ConnectPins(Optional ByVal site As Long = -1&)
    
    Call ConnectPins(m_PinList, site)

End Sub

'#Pass-Ok
Private Sub IVISVISource_DisconnectPins(Optional ByVal site As Long = -1&)
    
    Call DisconnectPins(m_PinList, site)

End Sub
'#V21-Release
Private Sub IVISVISource_DisconnectPinsMulti()
    
    Call DisconnectPinsMulti(m_PinList)

End Sub
'#Pass-Ok
Private Sub IVISVISource_ChangeMIRange(ByVal ClampI As Double, Optional ByVal site As Long = -1&)
    
    Call ChangeMIRange(m_PinList, CreateLimit(Abs(ClampI), 50 * mA, 1 * A), GetDpsMIRange(ClampI), site)

End Sub
'#V21-Release
Private Sub IVISVISource_ChangeMIRangeMulti(ByVal ClampI As Double)

    Call ChangeMIRangeMulti(m_PinList, CreateLimit(Abs(ClampI), 50 * mA, 1 * A), GetDpsMIRange(ClampI))

End Sub


'#Pass-Ok
Private Sub IVISVISource_MeasureV(retResult() As Double, ByVal avgNum As Long, Optional ByVal site As Long = -1&)
    
    Call OutputErrMsg(m_PinList & " is Invalid Pin Name at MeasureV()")
    Stop

End Sub

'#Pass-Ok
Private Sub IVISVISource_MeasureI(retResult() As Double, ByVal avgNum As Long, Optional ByVal site As Long = -1&)

    Call MeasureI(m_PinList, retResult, avgNum, site)

End Sub

'#Pass-Ok
Private Sub IVISVISource_MeasureVMulti(ByVal avgNum As Long)
        
    Call OutputErrMsg(m_PinList & " is Invalid Pin List at MeasureVMulti()")
    Stop

End Sub

'#Pass-Ok
Private Sub IVISVISource_MeasureIMulti(ByVal avgNum As Long)

    Call InitMultiDCResult
    Call MeasureIMulti(m_PinList, avgNum)

End Sub

'#Pass-Ok
Private Sub IVISVISource_ReadMVMultiResult(retResult() As Double)

    Call OutputErrMsg(m_PinList & " is Invalid Pin Name at ReadMVMultiResult()")
    Stop

End Sub

'#Pass-Ok
Private Sub IVISVISource_ReadMIMultiResult(retResult() As Double)
    
    Dim status As Boolean
    
    status = ReadMultiResult(m_PinList, retResult, m_ResultsI_DPS)

    If status = False Then
        Call OutputErrMsg("Not Found Result of " & m_PinList & " at ReadMIMultiResult()")
        Stop
    End If

End Sub

'################################################################################


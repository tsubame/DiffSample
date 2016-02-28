VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVISPpmu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   PPMUボード制御用クラス
'
'目的:
'   Ⅰ:PPMUボードの使い勝手向上
'   Ⅱ:V/ISourceとしてのI/F提供
'
'作成者:
'   SLSI今手
'
'XLibSTD_CommonDCModよりPPMU関連機能を整理
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
Private Const DEF_RANGE = -1
Private Const BOARD_NAME = "dcPPMU"
Private Const PPMU_VOLTAGE_RANGE = 2# * V

Private Const FV_LIMIT_HIGH = 7# * V
Private Const FV_LIMIT_LOW = -2# * V
Private Const FI_LIMIT_HIGH = 2# * mA
Private Const FI_LIMIT_LOW = -2# * mA

Private Const FV_RANGE_LIST = "None" '-2to7Vだが設定パラメータはない
Private Const FI_RANGE_LIST = "200uA,2mA"
Private Const MV_RANGE_LIST = "None" '2Vだが設定パラメータはない
Private Const MI_RANGE_LIST = "200nA,2uA,20uA,200uA,2mA"

Private m_PinList As String
Private m_ResultsV_PPMU As Collection
Private m_ResultsI_PPMU As Collection

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
    ByVal PinList As String, ByVal ForceV As Variant, ByVal MIRange As PpmuIRange, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'   PPMUを電圧印加状態に設定する｡
'
'パラメータ:
'    [PinList]    In   対象ピンリスト。
'    [ForceV]     In   印加電圧。配列指定可能。
'    [MIRange]    In   電流測定レンジ。
'    [Site]       In   サイト指定。オプション(Default -1)
'    [ConnectOn]  In   デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するPPMUを電圧印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ForceVで印加電圧を指定。ForceVは数値 or サイト数分の配列。
'    ■数値の場合全サイト同じ値を設定｡配列の場合各サイトに対応する配列要素の値を設定｡
'    ■MIRangeで測定レンジを指定｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    
    'Error Check ************************************
    Const FunctionName = "SetFVMI_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceV, -2 * V, 7 * V, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    Call ConvertVariableToArray(ForceV)

    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            TheHdw.PPMU.Pins(PinList).ForceVoltage(MIRange) = ForceV(curSite)
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
    ByVal PinList As String, ByVal ForceV As Variant, ByVal MIRange As PpmuIRange, _
     Optional ByVal ConnectOn As Boolean = True _
)

'内容:
'   PPMUを電圧印加状態に設定する｡ (Site同時)
'
'パラメータ:
'    [PinList]    In   対象ピンリスト。
'    [ForceV]     In   印加電圧。配列指定可能。
'    [MIRange]    In   電流測定レンジ。
'    [ConnectOn]  In   デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するPPMUを電圧印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ForceVで印加電圧を指定。ForceVは数値 or サイト数分の配列。
'    ■全サイト同じ値を設定｡
'    ■MIRangeで測定レンジを指定｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
'    Dim SiteStatus As loopStatus
    Dim curSite As Long
    Dim Channels() As Long
    
    curSite = 0
    'Error Check ************************************
    Const FunctionName = "SetFVMIMulti_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceV, -2 * V, 7 * V, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
    Call ConvertVariableToArray(ForceV)

         
            'Main Part ---------------------------------------
'            Pins指定
            TheHdw.PPMU.Pins(PinList).ForceVoltage(MIRange) = ForceV(curSite)
'            Chans指定
'            Call GetActiveChanList(PinList, chIO, Channels)
'            TheHdw.PPMU.Chans(Channels).ForceVoltage(MIRange) = ForceV(CurSite)
            '-------------------------------------------------
      
    
    If ConnectOn = True Then
        Call ConnectPinsMultiPins(PinList)
'        Call ConnectPinsMultiChans(pinList)
    End If

End Sub
'#Pass-Ok
Private Sub SetFIMV( _
    ByVal PinList As String, ByVal ForceI As Variant, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal FIRange As PpmuIRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'    PPMUを電流印加状態に設定する｡
'
'パラメータ:
'    [PinList]    In   対象ピンリスト。
'    [ForceI]     In   印加電流。配列指定可能。
'    [Site]       In   サイト指定。オプション(Default -1)
'    [FIRange]    In   電流印加レンジ。オプション(Default -1)
'    [ConnectOn]  In   デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するPPMUを電流印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ForceIで印加電流を指定。ForceIは数値 or サイト数分の配列。
'    ■数値の場合全サイト同じ値を設定｡配列の場合各サイトに対応する配列要素の値を設定｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■FIRangeで測定レンジを指定。-1(デフォルト)の場合、ForceIからレンジを判定して設定する。
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim FIRangeBySite() As PpmuIRange
    
    Dim Channels() As Long

    'Error Check ************************************
    Const FunctionName = "SetFIMV_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceI, -2 * mA, 2 * mA, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    Call ConvertVariableToArray(ForceI)

    ReDim FIRangeBySite(CountExistSite)
    For curSite = 0 To CountExistSite
        If FIRange = DEF_RANGE Then
            FIRangeBySite(curSite) = GetPpmuFIRange(ForceI(curSite))
        Else
            FIRangeBySite(curSite) = FIRange
        End If
    Next curSite
    
    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            TheHdw.PPMU.Pins(PinList).ForceCurrent(FIRangeBySite(curSite)) = ForceI(curSite)
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
Private Sub SetFIMVMulti( _
    ByVal PinList As String, ByVal ForceI As Variant, _
    Optional ByVal FIRange As PpmuIRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)

'内容:
'    PPMUを電流印加状態に設定する｡(Site同時)
'
'パラメータ:
'    [PinList]    In   対象ピンリスト。
'    [ForceI]     In   印加電流。配列指定可能。
'    [FIRange]    In   電流印加レンジ。オプション(Default -1)
'    [ConnectOn]  In   デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するPPMUを電流印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ForceIで印加電流を指定。ForceIは数値 or サイト数分の配列。
'    ■全サイト同じ値を設定｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■FIRangeで測定レンジを指定。-1(デフォルト)の場合、ForceIからレンジを判定して設定する。
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
'    Dim SiteStatus As loopStatus
    Dim curSite As Long
    Dim FIRangeBySite() As PpmuIRange
    
    Dim Channels() As Long

    'Error Check ************************************
    Const FunctionName = "SetFIMVMulti_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceI, -2 * mA, 2 * mA, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
    'site0only設定************************************
    Call ConvertVariableToArray(ForceI)

    ReDim FIRangeBySite(1)
   curSite = 0
        If FIRange = DEF_RANGE Then
            FIRangeBySite(curSite) = GetPpmuFIRange(ForceI(curSite))
        Else
            FIRangeBySite(curSite) = FIRange
        End If
    '************************************************
    
        
            'Main Part ---------------------------------------
'            Pins指定
            TheHdw.PPMU.Pins(PinList).ForceCurrent(FIRangeBySite(curSite)) = ForceI(curSite)
'            Chans指定
'            Call GetActiveChanList(PinList, chIO, Channels)
'            TheHdw.PPMU.Chans(Channels).ForceCurrent(FIRangeBySite(CurSite)) = ForceI(CurSite)
            '-------------------------------------------------
        
    
        If ConnectOn = True Then
        Call ConnectPinsMultiPins(PinList)
'        Call ConnectPinsMultiChans(pinList)
    End If
    
End Sub
'#Pass-Ok
Private Sub SetMV(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE, Optional ConnectOn As Boolean = True)
'内容:
'    PPMUを0A印加状態に設定する｡
'
'パラメータ:
'    [PinList]     In    対象ピンリスト。
'    [Site]        In    サイト指定。オプション(Default -1)
'    [ConnectOn]   In    デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するPPMUを0A印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
    Call SetFIMV(PinList, 0 * A, site, ppmu200uA, ConnectOn)

End Sub
'#V21-Release
Private Sub SetMVMulti(ByVal PinList As String, Optional ConnectOn As Boolean = True)
'内容:
'    PPMUを0A印加状態に設定する｡(Site同時)
'
'パラメータ:
'    [PinList]     In    対象ピンリスト。
'    [ConnectOn]   In    デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するPPMUを0A印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
    Call SetFIMVMulti(PinList, 0 * A, ppmu200uA, ConnectOn)
End Sub
'#Pass-Ok
Private Sub DisconnectPins(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)
'内容:
'   PPMUをデバイスから切り離す｡
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'    [Site]      In   サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するPPMUをデバイスから切り離す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■切り離した後、0V印加状態に設定する(接続はしない)。
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    
    'Error Check ************************************
    Const FunctionName = "DisconnectPins_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************

    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            TheHdw.Digital.relays.Pins(PinList).DisconnectPins
            
            With TheHdw.PPMU.Pins(PinList)
                .ForceCurrent(ppmu2mA) = 0 * A
                .ForceVoltage(ppmuAutoRange) = 0 * V
            End With
            '-------------------------------------------------
            
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

End Sub
'#V21-Release
Private Sub DisconnectPinsMulti(ByVal PinList As String)
'内容:
'   PPMUをデバイスから切り離す｡(Site同時)
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するPPMUをデバイスから切り離す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡ (API機能)
'    ■切り離した後、0V印加状態に設定する(接続はしない)。
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long
    
    'Error Check ************************************
    Const FunctionName = "DisconnectPinsMulti_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
        
    'Main Part ---------------------------------------
    TheHdw.Digital.relays.Pins(PinList).DisconnectPins
            
    With TheHdw.PPMU.Pins(PinList)
    .ForceCurrent(ppmu2mA) = 0 * A
    .ForceVoltage(ppmuAutoRange) = 0 * V
    End With
    '-------------------------------------------------
            
End Sub
'#Pass-Ok
Private Sub ConnectPins(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)
'内容:
'   PPMUをデバイスに接続する｡
'
'パラメータ:
'    [PinList]   In    対象ピンリスト。
'    [Site]      In    サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するPPMUをデバイスに接続する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    
    'Error Check ************************************
    Const FunctionName = "ConnectPins_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************

    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            TheHdw.PPMU.Pins(PinList).Connect
            '-------------------------------------------------
            
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

End Sub
'#V21-Release
Private Sub ConnectPinsMultiPins(ByVal PinList As String)
'内容:
'   PPMUをデバイスに接続する｡(site同時)
'
'パラメータ:
'    [PinList]   In    対象ピンリスト。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するPPMUをデバイスに接続する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long
    
    'Error Check ************************************
    Const FunctionName = "ConnectPinsMulti_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************

    'Main *************************************
            'Main Part ---------------------------------------
            TheHdw.PPMU.Pins(PinList).Connect
            '-------------------------------------------------
    '************************************************

End Sub
'#V21-Release
Private Sub ConnectPinsMultiChans(ByVal PinList As String)
'内容:
'   PPMUをデバイスに接続する｡(site同時)
'
'パラメータ:
'    [PinList]   In    対象ピンリスト。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するPPMUをデバイスに接続する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long
     Dim Channels() As Long
    
    'Error Check ************************************
    Const FunctionName = "ConnectPinsMulti_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************

  
            'Main Part ---------------------------------------
            Call GetActiveChanList(PinList, chIO, Channels)
            TheHdw.PPMU.Chans(Channels).Connect
            '-------------------------------------------------


End Sub
'#Pass-Ok
Private Sub SetGND( _
    ByVal PinList As String, Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'   PPMUを0V印加状態に設定する｡
'
'パラメータ:
'    [PinList]      In    対象ピンリスト。
'    [Site]         In    サイト指定。オプション(Default -1)
'    [ConnectOn]    In    デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するPPMUを0V印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
    Call SetFVMI(PinList, 0 * V, ppmu2mA, site, ConnectOn)
    
End Sub

'#Pass-Ok
Private Sub ChangeMIRange(ByVal PinList As String, ByVal MIRange As PpmuIRange, Optional ByVal site As Long = ALL_SITE)
'内容:
'   電圧印加状態のPPMUの電流測定レンジを変更する｡
'
'パラメータ:
'    [PinList]    In    対象ピンリスト。
'    [MIRange]    In    電流測定レンジ。
'    [Site]       In    サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するPPMUの印加電圧を変えず､電流測定レンジを変更する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■MIRangeで測定レンジを指定｡選択肢が表示される｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long
    Dim ForceV As Double
    
    Dim ErrChans As String
    
    'Error Check ************************************
    Const FunctionName = "ChangeMIRange_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    'Site Loop **************************************
    ErrChans = ""
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            Call GetChanList(PinList, curSite, chIO, Channels)
            For chan = 0 To UBound(Channels)
                With TheHdw.PPMU.Chans(Channels(chan))
                    If .IsForcingVoltage = True Then
                        ForceV = .ForceVoltage(.CurrentRange)
                        .ForceVoltage(MIRange) = ForceV
                    Else
                        ErrChans = ErrChans & Channels(chan) & ","
'                        Call OutputErrMsg("ch" & Channels(chan) & " is not MI Mode.")
                    End If
                End With
            Next chan
            '-------------------------------------------------
            
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

    If ErrChans <> "" Then
        Call OutputErrMsg("PPMU(" & ErrChans & ") is not MI Mode. (at ChangeMIRange_PPMU)")
    End If

End Sub

'#V21-Release
Private Sub ChangeMIRangeMulti( _
    ByVal PinList As String, ByVal MIRange As PpmuIRange _
   )
'内容:
'   電圧印加状態のPPMUの電流測定レンジを変更する｡(Site同時)
'
'パラメータ:
'    [PinList]    In    対象ピンリスト。
'    [MIRange]    In    電流測定レンジ。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するPPMUの印加電圧を変えず､電流測定レンジを変更する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■MIRangeで測定レンジを指定｡選択肢が表示される｡
'    ■非アクティブサイトに対しては何もしない｡
'
'    Dim SiteStatus As loopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long
    Dim ForceV As Double
    
    Dim ErrChans As String
    
    'Error Check ************************************
    Const FunctionName = "ChangeMIRangeMulti_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
    ErrChans = ""
        
            'Main Part ---------------------------------------
            Call GetActiveChanList(PinList, chIO, Channels)
 
                With TheHdw.PPMU.Chans(Channels)
                    If .IsForcingVoltage = True Then
                        ForceV = .ForceVoltage(.CurrentRange)
                        .ForceVoltage(MIRange) = ForceV
                    Else
                        For chan = 0 To UBound(Channels)
                        ErrChans = ErrChans & Channels(chan) & ","
                        Next chan
'                        Call OutputErrMsg("ch" & Channels(chan) & " is not MI Mode.")
                    End If
                End With
            '-------------------------------------------------
            
 

    If ErrChans <> "" Then
        Call OutputErrMsg("PPMU(" & ErrChans & ") is not MI Mode. (at ChangeMIRange_PPMU)")
    End If
    
End Sub
'#Pass-Ok
Private Sub MeasureV( _
    ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE _
)
'内容:
'   PPMUの電圧値を測定する｡
'
'パラメータ:
'    [PinName]       In    対象ピン名。
'    [retResult()]   Out   結果を格納する配列。(サイト数分の要素が必要)
'    [AvgNum]        In    平均回数。
'    [Site]          In    サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinNameに対応するPPMUの電圧値を測定する｡
'    ■PinNameは1ピンのみ｡複数ピン､ピングループの指定は不可｡
'    ■retResultに結果を返す｡
'    ■AvgNumの数だけサンプリングして､その平均を返す｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■Siteが-1以外の場合、対象サイト以外のretResult()の値は変化させない。
'    ■非アクティブサイトの結果は0を返す。Siteが-1の時も同様。
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long
    Dim Samples() As Double
    Dim n As Long
    
    Dim ErrChans As String
    
    'Error Check ************************************
    Const FunctionName = "MeasureV_PPMU"
    If CheckPinList(PinName, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckSinglePins(PinName, chIO, FunctionName) = False Then Stop: Exit Sub
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
    ErrChans = ""
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            Call GetChanList(PinName, curSite, chIO, Channels)
            
            With TheHdw.PPMU.Chans(Channels(0))
                If .isForcingCurrent = False Then
                    ErrChans = ErrChans & Channels(chan) & ","
'                    Call OutputErrMsg(PinName & " is not MV Mode (at MeasureV_PPMU)")
                Else
                    TheHdw.PPMU.Samples = avgNum
                    Call .MeasureVoltages(Samples)
                    TheHdw.PPMU.Samples = 1
                    
                    retResult(curSite) = 0
                    For n = 0 To UBound(Samples)
                        retResult(curSite) = retResult(curSite) + Samples(n)
                    Next n
                    retResult(curSite) = retResult(curSite) / avgNum
                End If
            End With
            '-------------------------------------------------
            
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

    If ErrChans <> "" Then
        Call OutputErrMsg("PPMU(" & ErrChans & ") is not MV Mode (at MeasureV_PPMU)")
    End If
    
End Sub

'#Pass-Ok
Private Sub MeasureI( _
    ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE _
)
'内容:
'   PPMUの電流値を測定する｡
'
'パラメータ:
'    [PinName]         In    対象ピン名。
'    [retResult()]     Out   結果を格納する配列。(サイト数分の要素が必要)
'    [AvgNum]          In    平均回数。
'    [Site]            In    サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinNameに対応するPPMUの電流値を測定する｡
'    ■PinNameは1ピンのみ｡複数ピン､ピングループの指定は不可｡
'    ■retResultに結果を返す｡
'    ■AvgNumの数だけサンプリングして､その平均を返す｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■Siteが-1以外の場合、対象サイト以外のretResult()の値は変化させない。
'    ■非アクティブサイトの結果は0を返す。Siteが-1の時も同様。
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long
    Dim Samples() As Double
    Dim n As Long
    
    Dim ErrChans As String
    
    'Error Check ************************************
    Const FunctionName = "MeasureI_PPMU"
    If CheckPinList(PinName, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckSinglePins(PinName, chIO, FunctionName) = False Then Stop: Exit Sub
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
    ErrChans = ""
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            Call GetChanList(PinName, curSite, chIO, Channels)
            
            With TheHdw.PPMU.Chans(Channels(0))
                If .IsForcingVoltage = False Then
                    ErrChans = ErrChans & Channels(chan) & ","
'                    Call OutputErrMsg(PinName & " is not MI Mode (at MeasureI_PPMU)")
                Else
                    TheHdw.PPMU.Samples = avgNum
                    Call .MeasureCurrents(Samples)
                    TheHdw.PPMU.Samples = 1
                    
                    retResult(curSite) = 0
                    For n = 0 To UBound(Samples)
                        retResult(curSite) = retResult(curSite) + Samples(n)
                    Next n
                    retResult(curSite) = retResult(curSite) / avgNum
                End If
            End With
            '-------------------------------------------------
            
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

    If ErrChans <> "" Then
        Call OutputErrMsg("PPMU(" & ErrChans & ") is not MI Mode (at MeasureI_PPMU)")
    End If

End Sub

'#Pass-Ok
Private Sub MeasureVMulti(ByVal PinList As String, ByVal avgNum As Long)
'内容:
'   PPMUの電圧値を測定する｡
'
'パラメータ:
'    [PinList]     In    対象ピン名。
'    [AvgNum]      In    平均回数。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するピンの電圧値を測定する｡結果はReadMVMultiResultを使って読み出す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
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
    Const FunctionName = "MeasureVMulti_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckAvgNum(avgNum, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    'Measurement **************************************************
    Call GetActiveChanList(PinList, chIO, Channels)
    With TheHdw.PPMU
        .Samples = avgNum
        Call .Chans(Channels).MeasureVoltages(Samples)
        .Samples = 1
    End With
    '**************************************************************
    
    'Regist Results ***********************************************
    'Initialize ----------------------------------------------
    Call SeparatePinList(PinList, pinNames)
    PerPinResults = CreateEmpty2DArray(UBound(pinNames), CountExistSite)
    '---------------------------------------------------------
    
    'Summation -----------------------------------------------
    i = 0
    For n = 1 To avgNum
        For Pin = 0 To UBound(pinNames)
            For curSite = 0 To CountExistSite
                If IsActiveSite(curSite) Then
                    PerPinResults(Pin)(curSite) = PerPinResults(Pin)(curSite) + Samples(i)
                    i = i + 1
                End If
            Next curSite
        Next Pin
    Next n
    '---------------------------------------------------------
    
    'Average & Regist ----------------------------------------
    Set m_ResultsV_PPMU = New Collection
    For Pin = 0 To UBound(pinNames)
        For curSite = 0 To CountExistSite
            PerPinResults(Pin)(curSite) = PerPinResults(Pin)(curSite) / avgNum
        Next curSite
        Call m_ResultsV_PPMU.Add(PerPinResults(Pin), pinNames(Pin))
    Next Pin
    '---------------------------------------------------------
    '**************************************************************
    
End Sub

'#Pass-Ok
Private Sub MeasureIMulti(ByVal PinList As String, ByVal avgNum As Long)
'内容:
'   PPMUの電流値を測定する｡
'
'パラメータ:
'    [PinList]     In    対象ピン名。
'    [AvgNum]      In    平均回数。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するピンの電流値を測定する｡結果はReadMIMultiResultを使って読み出す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
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
    Const FunctionName = "MeasureIMulti_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckAvgNum(avgNum, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    'Measurement **************************************************
    Call GetActiveChanList(PinList, chIO, Channels)
    With TheHdw.PPMU
        .Samples = avgNum
        Call .Chans(Channels).MeasureCurrents(Samples)
        .Samples = 1
    End With
    '**************************************************************
    
    'Regist Results ***********************************************
    'Initialize ----------------------------------------------
    Call SeparatePinList(PinList, pinNames)
    PerPinResults = CreateEmpty2DArray(UBound(pinNames), CountExistSite)
    '---------------------------------------------------------
    
    'Summation -----------------------------------------------
    i = 0
    For n = 1 To avgNum
        For Pin = 0 To UBound(pinNames)
            For curSite = 0 To CountExistSite
                If IsActiveSite(curSite) Then
                    PerPinResults(Pin)(curSite) = PerPinResults(Pin)(curSite) + Samples(i)
                    i = i + 1
                End If
            Next curSite
        Next Pin
    Next n
    '---------------------------------------------------------
    
    'Average & Regist ----------------------------------------
    Set m_ResultsI_PPMU = New Collection
    For Pin = 0 To UBound(pinNames)
        For curSite = 0 To CountExistSite
            PerPinResults(Pin)(curSite) = PerPinResults(Pin)(curSite) / avgNum
        Next curSite
        Call m_ResultsI_PPMU.Add(PerPinResults(Pin), pinNames(Pin))
    Next Pin
    '---------------------------------------------------------
    '**************************************************************
    
End Sub

'#Pass-Ok
Private Sub ReadMVMultiResult(ByVal PinName As String, ByRef retResult() As Double)
'内容:
'   MeasureVMultiで得た電圧値を読み出す｡
'
'パラメータ:
'    [PinName]        In   対象ピン名。
'    [retResult()]    Out  結果を格納する配列。(サイト数分の要素が必要)
'
'戻り値:
'
'注意事項
'    ■PinNameは最後に使ったMeasureVMultiのピンリストに含まれるピン名を指定｡
'    ■retResultに結果を返す｡
'
    Dim status As Boolean

    status = ReadMultiResult(PinName, retResult, m_ResultsV_PPMU)

    If status = False Then
        Call OutputErrMsg("Not Found Result of " & PinName & " at ReadMVMultiResult()")
        Stop
    End If

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
    
    status = ReadMultiResult(PinName, retResult, m_ResultsI_PPMU)

    If status = False Then
        Call OutputErrMsg("Not Found Result of " & PinName & " at ReadMIMultiResult()")
        Stop
    End If
    
End Sub

'#Pass-Ok
Private Sub InitMultiDCResult()
'内容:
'   Measure＊Mulitiで使用するコレクションを初期化
'
    Set m_ResultsV_PPMU = Nothing
    Set m_ResultsI_PPMU = Nothing

End Sub

'#Pass-Ok
Private Function GetPpmuFIRange(ByVal ForceI As Double) As PpmuIRange

    Dim AbsForceI As Double

    AbsForceI = Abs(ForceI)
    If AbsForceI <= 0.0002 Then         '200uA
        GetPpmuFIRange = ppmu200uA
    Else
        GetPpmuFIRange = ppmu2mA
    End If

End Function

'#Pass-Ok
Private Function GetPpmuMIRange(ByVal ClampI As Double) As PpmuIRange

    Dim AbsClampI As Double

    AbsClampI = Abs(ClampI)

    If AbsClampI <= 0.0000002 Then              '200nA
        GetPpmuMIRange = ppmu200nA
    ElseIf AbsClampI <= 0.000002 Then           '2uA
        GetPpmuMIRange = ppmu2uA
    ElseIf AbsClampI <= 0.00002 Then            '20uA
        GetPpmuMIRange = ppmu20uA
    ElseIf AbsClampI <= 0.0002 Then             '200uA
        GetPpmuMIRange = ppmu200uA
    Else
        GetPpmuMIRange = ppmu2mA
    End If

End Function

'#Pass-Ok
Private Function GetPpmuMeasureRange() As Collection
'PPMUの測定レンジを取得する
    
    Dim pinNameArr() As String
    Dim chListArr() As Long
    Dim pinLpCnt As Long
    Dim siteLpCnt As Long
    Dim tmpMeasureRng() As Double
    Dim tmpIrange() As PpmuIRange
    Dim tmpRngVal As Collection
                                
    Const FUNCTION_NAME = "GetPpmuMeasureRange"
                                                                
    '取得レンジ格納用変数
    ReDim tmpMeasureRng(CountExistSite)
    ReDim tmpIrange(CountExistSite)
    
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
            With TheHdw.PPMU.Chans(chListArr(siteLpCnt))
                If .IsForcingVoltage <> True Then
                    'FIMV
                    tmpMeasureRng(siteLpCnt) = PPMU_VOLTAGE_RANGE
                Else
                    'FVMI
                    tmpIrange(siteLpCnt) = .CurrentRange
                    tmpMeasureRng(siteLpCnt) = PpmuIRangeToValue(tmpIrange(siteLpCnt))
                End If
            End With
        Next siteLpCnt
        
        'ピン名をキーにレンジ値の取得結果をコレクションに登録
        Call tmpRngVal.Add(tmpMeasureRng, pinNameArr(pinLpCnt))
    
    Next pinLpCnt
    
    '取得結果を関数の戻り値としてコレクションで返す
    Set GetPpmuMeasureRange = tmpRngVal
        
End Function

'#Pass-Ok
Private Function GetPpmuForceValue() As Collection
'PPMUのForce値を取得する
    
    Dim pinNameArr() As String
    Dim chListArr() As Long
    Dim pinLpCnt As Long
    Dim siteLpCnt As Long
    Dim tmpForceValue() As Double
    Dim tmpForceVal As Collection
                                    
    Const FUNCTION_NAME = "GetPpmuForceValue"
                                                                
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
            With TheHdw.PPMU.Chans(chListArr(siteLpCnt))
                If .IsForcingVoltage <> True Then
                    'FIMV
                    tmpForceValue(siteLpCnt) = .ForceCurrent(.CurrentRange)
                Else
                    'FVMI
                    tmpForceValue(siteLpCnt) = .ForceVoltage(.CurrentRange)
                End If
            End With
        Next siteLpCnt
        'ピン名をキーにレンジ値の取得結果をコレクションに登録
        Call tmpForceVal.Add(tmpForceValue, pinNameArr(pinLpCnt))
    Next pinLpCnt
    
    '取得結果を関数の戻り値としてコレクションで返す
    Set GetPpmuForceValue = tmpForceVal
        
End Function

'#Pass-Ok
Private Function PpmuIRangeToValue(ByVal Irange As PpmuIRange) As Double
'PPMUのIRange番号をDouble値に変換

    Select Case Irange
        Case ppmuSmartRange
            PpmuIRangeToValue = DC_RANGE_SMART     '"Smart"
        
        Case ppmuAutoRange
            PpmuIRangeToValue = DC_RANGE_AUTO      '"Auto"
        
        Case Else
            PpmuIRangeToValue = TheHdw.PPMU.CurrentRangeToValue(Irange)
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

    Set IVISVISource_GetForceValue = GetPpmuForceValue
    
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
    
    Dim tmpLimitVal(1) As Double

    tmpLimitVal(0) = FI_LIMIT_LOW
    tmpLimitVal(1) = FI_LIMIT_HIGH

    IVISVISource_GetForceILimit = tmpLimitVal

End Function

'#Pass-Ok
Private Function IVISVISource_GetMeasureRange() As Collection
    
    Set IVISVISource_GetMeasureRange = Nothing
        
    'FAILサイトが存在するか確認し、存在する場合はエラーとする
    If CheckFailSiteExists("GetMeasureRange") = True Then
        Exit Function
    End If

    Set IVISVISource_GetMeasureRange = GetPpmuMeasureRange

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
    
    m_PinList = PinList

End Property

'#Pass-Ok
Private Sub IVISVISource_SetFVMI(ByVal ForceV As Variant, ByVal ClampI As Double, Optional ByVal site As Long = -1&, Optional ByVal ConnectOn As Boolean = True)

    Call SetFVMI(m_PinList, ForceV, GetPpmuMIRange(ClampI), site, ConnectOn)

End Sub
'#V21-Release
Private Sub IVISVISource_SetFVMIMulti(ByVal ForceV As Variant, ByVal ClampI As Double, Optional ByVal ConnectOn As Boolean = True)

    Call SetFVMIMulti(m_PinList, ForceV, GetPpmuMIRange(ClampI), ConnectOn)

End Sub
'#Pass-Ok
Private Sub IVISVISource_SetFIMV(ByVal ForceI As Variant, ByVal ClampV As Double, Optional ByVal site As Long = -1&, Optional ByVal ConnectOn As Boolean = True)

    Call SetFIMV(m_PinList, ForceI, site, , ConnectOn)

End Sub
'#V21-Release
Private Sub IVISVISource_SetFIMVMulti(ByVal ForceI As Variant, ByVal ClampV As Double, Optional ByVal ConnectOn As Boolean = True)

    Call SetFIMVMulti(m_PinList, ForceI, , ConnectOn)

End Sub
'#Pass-Ok
Private Sub IVISVISource_SetMV(ByVal ClampV As Double, Optional ByVal site As Long = -1&, Optional ByVal ConnectOn As Boolean = True)
    
    Call SetMV(m_PinList, site, ConnectOn)

End Sub
'#V21-Release
Private Sub IVISVISource_SetMVMulti(ByVal ClampV As Double, Optional ByVal ConnectOn As Boolean = True)
    
    Call SetMVMulti(m_PinList, ConnectOn)

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

    Call ChangeMIRange(m_PinList, GetPpmuMIRange(ClampI), site)

End Sub
'#V21-Release
Private Sub IVISVISource_ChangeMIRangeMulti(ByVal ClampI As Double)

    Call ChangeMIRangeMulti(m_PinList, GetPpmuMIRange(ClampI))

End Sub

'#Pass-Ok
Private Sub IVISVISource_MeasureV(retResult() As Double, ByVal avgNum As Long, Optional ByVal site As Long = -1&)

    Call MeasureV(m_PinList, retResult, avgNum, site)

End Sub

'#Pass-Ok
Private Sub IVISVISource_MeasureI(retResult() As Double, ByVal avgNum As Long, Optional ByVal site As Long = -1&)

    Call MeasureI(m_PinList, retResult, avgNum, site)

End Sub

'#Pass-Ok
Private Sub IVISVISource_MeasureVMulti(ByVal avgNum As Long)
    
    Call InitMultiDCResult
    Call MeasureVMulti(m_PinList, avgNum)

End Sub

'#Pass-Ok
Private Sub IVISVISource_MeasureIMulti(ByVal avgNum As Long)

    Call InitMultiDCResult
    Call MeasureIMulti(m_PinList, avgNum)

End Sub

'#Pass-Ok
Private Sub IVISVISource_ReadMVMultiResult(retResult() As Double)
    
    Dim status As Boolean

    status = ReadMultiResult(m_PinList, retResult, m_ResultsV_PPMU)

    If status = False Then
        Call OutputErrMsg("Not Found Result of " & m_PinList & " at ReadMVMultiResult()")
        Stop
    End If

End Sub

'#Pass-Ok
Private Sub IVISVISource_ReadMIMultiResult(retResult() As Double)

    Dim status As Boolean

    status = ReadMultiResult(m_PinList, retResult, m_ResultsI_PPMU)

    If status = False Then
        Call OutputErrMsg("Not Found Result of " & m_PinList & " at ReadMIMultiResult()")
        Stop
    End If

End Sub

'################################################################################


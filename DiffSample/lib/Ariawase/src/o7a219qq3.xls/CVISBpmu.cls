VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVISBpmu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   BPMUボード制御用クラス
'
'目的:
'   Ⅰ:BPMUボード使い勝手向上
'   Ⅱ:V/ISourceとしてのI/F提供
'
'作成者:
'   SLSI今手
'
'XLibSTD_CommonDCModよりBPMU関連機能を整理
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
Private Const BOARD_NAME = "dcBPMU"

Private Const FV_LIMIT_HIGH = 24# * V
Private Const FV_LIMIT_LOW = -24# * V
Private Const FI_LIMIT_HIGH = 200# * mA
Private Const FI_LIMIT_LOW = -200# * mA

Private Const FV_RANGE_LIST = "2V,5V,10V,24V"
Private Const MV_RANGE_LIST = "2V,5V,10V,24V"
Private Const FI_RANGE_LIST = "200uA,2mA,20mA,200mA"
Private Const MI_RANGE_LIST = "20uA,200uA,2mA,20mA,200mA"

'ボードの機能としてMIレンジに2uA設定は存在するがクランプ機能が
'無効となるので電源クラスでは選択されないようにしている
'Private Const MI_RANGE_LIST = "2uA,20uA,200uA,2mA,20mA,200mA"

Private m_PinList As String

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
    Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal FVRange As BpmuVRange = DEF_RANGE, Optional ByVal MIRange As BpmuIRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'   BPMUを電圧印加状態に設定する｡
'
'パラメータ:
'    [PinList]    In   対象ピンリスト。
'    [ForceV]     In   印加電圧。配列指定可能。
'    [ClampI]     In   クランプ電流値。
'    [Site]       In   サイト指定。オプション(Default -1)
'    [FVRange]    In   電圧印加レンジ。オプション(Default -1)
'    [MIRange]    In   電流測定レンジ。オプション(Default -1)
'    [ConnectOn]  In   デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■MIRangeがbpmu2uA設定のときは電流クランプ機能が働きません。
'    ■PinListに対応するBPMUを電圧印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ForceVで印加電圧を指定。ForceVは数値 or サイト数分の配列。
'    ■数値の場合全サイト同じ値を設定｡配列の場合各サイトに対応する配列要素の値を設定｡
'    ■ClampIでクランプ電流を設定｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■FVRangeで印加電圧レンジを設定。-1(デフォルト)の場合、ForceVからレンジを判定して設定する。
'    ■MIRangeで測定レンジを指定。-1(デフォルト)の場合、ClampIからレンジを判定して設定する。
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim FVRangeBySite() As BpmuVRange
    
    'Error Check ************************************
    Const FunctionName = "SetFVMI_BPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceV, -24 * V, 24 * V, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 0 * mA, 200 * mA, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    Call ConvertVariableToArray(ForceV)
    
    ReDim FVRangeBySite(CountExistSite)
    For curSite = 0 To CountExistSite
        If FVRange = DEF_RANGE Then
            FVRangeBySite(curSite) = GetBpmuFVRange(ForceV(curSite))
        Else
            FVRangeBySite(curSite) = FVRange
        End If
    Next curSite
    
    If MIRange = DEF_RANGE Then
        MIRange = GetBpmuMIRange(ClampI)
    End If
    
    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            With TheHdw.BPMU.Pins(PinList)
                .ClampCurrent(MIRange) = ClampI
                .ForceVoltage(FVRangeBySite(curSite)) = ForceV(curSite)
                Call .ModeFVMI(MIRange, FVRangeBySite(curSite))
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
Private Sub SetFVMIMulti(ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, _
    Optional ByVal FVRange As BpmuVRange = DEF_RANGE, Optional ByVal MIRange As BpmuIRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
    )

'内容:
'   BPMUを電圧印加状態に設定する｡ (Site同時)
'
'パラメータ:
'    [PinList]    In   対象ピンリスト。
'    [ForceV]     In   印加電圧。配列指定可能。
'    [ClampI]     In   クランプ電流値。
'    [FVRange]    In   電圧印加レンジ。オプション(Default -1)
'    [MIRange]    In   電流測定レンジ。オプション(Default -1)
'    [ConnectOn]  In   デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■MIRangeがbpmu2uA設定のときは電流クランプ機能が働きません。
'    ■PinListに対応するBPMUを電圧印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ForceVで印加電圧を指定。ForceVは数値 or サイト数分の配列。
'    ■全サイト同じ値を設定｡
'    ■ClampIでクランプ電流を設定｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■FVRangeで印加電圧レンジを設定。-1(デフォルト)の場合、ForceVからレンジを判定して設定する。
'    ■MIRangeで測定レンジを指定。-1(デフォルト)の場合、ClampIからレンジを判定して設定する。
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
'    Dim SiteStatus As loopStatus
    Dim curSite As Long
    Dim FVRangeBySite() As BpmuVRange
    Dim Channels() As Long
    
    'Error Check ************************************
    Const FunctionName = "SetFVMIMulti_BPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceV, -24 * V, 24 * V, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 0 * mA, 200 * mA, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
'site0only設定************************************
    Call ConvertVariableToArray(ForceV)
    
        ReDim FVRangeBySite(1)
        curSite = 0
        If FVRange = DEF_RANGE Then
            FVRangeBySite(curSite) = GetBpmuFVRange(ForceV(curSite))
        Else
            FVRangeBySite(curSite) = FVRange
        End If
     '************************************************
    
    
    If MIRange = DEF_RANGE Then
        MIRange = GetBpmuMIRange(ClampI)
    End If
    
    
            'Main Part ---------------------------------------
'Pins指定
            With TheHdw.BPMU.Pins(PinList)
                .ClampCurrent(MIRange) = ClampI
                .ForceVoltage(FVRangeBySite(curSite)) = ForceV(curSite)
                Call .ModeFVMI(MIRange, FVRangeBySite(curSite))
            End With
' Chans指定
'            Call GetActiveChanList(pinList, chIO, Channels)
'            With TheHdw.BPMU.Chans(Channels)
'                .ClampCurrent(MIRange) = ClampI
'                .ForceVoltage(FVRangeBySite(CurSite)) = ForceV(CurSite)
'                Call .ModeFVMI(MIRange, FVRangeBySite(CurSite))
'            End With
            '-------------------------------------------------
            


    If ConnectOn = True Then
        Call ConnectPinsMultiPins(PinList)
'       Call ConnectPinsMultiChans(pinList)
    End If
 


End Sub
'#Pass-Ok
Private Sub SetFIMV( _
    ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, _
    Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal FIRange As BpmuIRange = DEF_RANGE, Optional ByVal MVRange As BpmuVRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'   BPMUを電流印加状態に設定する｡
'
'パラメータ:
'    [PinList]    In   対象ピンリスト。
'    [ForceI]     In   印加電流。配列指定可能。
'    [ClampV]     In   クランプ電圧値。
'    [Site]       In   サイト指定。オプション(Default -1)
'    [FIRange]    In   電流印加レンジ。オプション(Default -1)
'    [MVRange]    In   電圧測定レンジ。オプション(Default -1)
'    [ConnectOn]  In   デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するBPMUを電流印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ForceIで印加電流を指定。ForceIは数値 or サイト数分の配列。
'    ■数値の場合全サイト同じ値を設定｡配列の場合各サイトに対応する配列要素の値を設定｡
'    ■ClampVでクランプ電圧を設定｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■FIRangeで印加電流レンジを設定。-1(デフォルト)の場合、ForceIからレンジを判定して設定する。
'    ■MVRangeで測定レンジを指定。-1(デフォルト)の場合、ClampVからレンジを判定して設定する。
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim FIRangeBySite() As BpmuIRange

    'Error Check ************************************
    Const FunctionName = "SetFIMV_BPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceI, -200 * mA, 200 * mA, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampV, 0 * V, 24 * V, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    Call ConvertVariableToArray(ForceI)

    ReDim FIRangeBySite(CountExistSite)
    For curSite = 0 To CountExistSite
        If FIRange = DEF_RANGE Then
            FIRangeBySite(curSite) = GetBpmuFIRange(ForceI(curSite))
        Else
            FIRangeBySite(curSite) = FIRange
        End If
    Next curSite
    
    If MVRange = DEF_RANGE Then
        MVRange = GetBpmuMVRange(ClampV)
    End If
    
    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then

            'Main Part ---------------------------------------
            With TheHdw.BPMU.Pins(PinList)
                .ClampVoltage(MVRange) = ClampV
                .ForceCurrent(FIRangeBySite(curSite)) = ForceI(curSite)
                Call .ModeFIMV(FIRangeBySite(curSite), MVRange)
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
Private Sub SetFIMVMulti( _
  ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, _
    Optional ByVal FIRange As BpmuIRange = DEF_RANGE, Optional ByVal MVRange As BpmuVRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)

'内容:
'   BPMUを電流印加状態に設定する｡ (Site同時)
'
'パラメータ:
'    [PinList]    In   対象ピンリスト。
'    [ForceI]     In   印加電流。配列指定可能。
'    [ClampV]     In   クランプ電圧値。
'    [FIRange]    In   電流印加レンジ。オプション(Default -1)
'    [MVRange]    In   電圧測定レンジ。オプション(Default -1)
'    [ConnectOn]  In   デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するBPMUを電流印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ForceIで印加電流を指定。ForceIは数値 or サイト数分の配列。
'    ■全サイト同じ値を設定｡
'    ■ClampVでクランプ電圧を設定｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■FIRangeで印加電流レンジを設定。-1(デフォルト)の場合、ForceIからレンジを判定して設定する。
'    ■MVRangeで測定レンジを指定。-1(デフォルト)の場合、ClampVからレンジを判定して設定する。
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
'    Dim SiteStatus As loopStatus
    Dim curSite As Long
    Dim FIRangeBySite() As BpmuIRange
    Dim Channels() As Long
    
    'Error Check ************************************
    Const FunctionName = "SetFIMVMulti_BPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceI, -200 * mA, 200 * mA, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampV, 0 * V, 24 * V, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
    'site0only設定************************************
    Call ConvertVariableToArray(ForceI)

    ReDim FIRangeBySite(1)
    curSite = 0
        If FIRange = DEF_RANGE Then
            FIRangeBySite(curSite) = GetBpmuFIRange(ForceI(curSite))
        Else
            FIRangeBySite(curSite) = FIRange
        End If
   '************************************************
    
    If MVRange = DEF_RANGE Then
        MVRange = GetBpmuMVRange(ClampV)
    End If
    
   
            'Main Part ---------------------------------------
'Pins指定
            With TheHdw.BPMU.Pins(PinList)
                .ClampVoltage(MVRange) = ClampV
                .ForceCurrent(FIRangeBySite(curSite)) = ForceI(curSite)
                Call .ModeFIMV(FIRangeBySite(curSite), MVRange)
            End With
'Chans指定
'            Call GetActiveChanList(pinList, chIO, Channels)
'            With TheHdw.BPMU.Chans(Channels)
'                .ClampVoltage(MVRange) = ClampV
'                .ForceCurrent(FIRangeBySite(CurSite)) = ForceI(CurSite)
'                Call .ModeFIMV(FIRangeBySite(CurSite), MVRange)
'            End With
            '-------------------------------------------------

    If ConnectOn = True Then
        Call ConnectPinsMultiPins(PinList)
'        Call ConnectPinsMultiChans(pinList)
    End If
    
End Sub
'#Pass-Ok
Private Sub SetMV( _
    ByVal PinList As String, ByVal ClampV As Double, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal MVRange As BpmuVRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'   BPMUを0A印加状態に設定する｡
'
'パラメータ:
'    [PinList]    In   対象ピンリスト。
'    [ClampV]     In   クランプ電圧値。
'    [Site]       In   サイト指定。オプション(Default -1)
'    [MVRange]    In   電圧測定レンジ。オプション(Default -1)
'    [ConnectOn]  In   デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するBPMUを0A印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ClampVでクランプ電圧を設定｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■MVRangeで測定レンジを指定。-1(デフォルト)の場合、ClampVからレンジを判定して設定する。
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
    Call SetFIMV(PinList, 0 * A, ClampV, site, , MVRange, ConnectOn)

End Sub
'#V21-Release
Private Sub SetMVMulti(ByVal PinList As String, ByVal ClampV As Double, _
    Optional ByVal MVRange As BpmuVRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'   BPMUを0A印加状態に設定する｡(Site同時)
'
'パラメータ:
'    [PinList]    In   対象ピンリスト。
'    [ClampV]     In   クランプ電圧値。
'    [MVRange]    In   電圧測定レンジ。オプション(Default -1)
'    [ConnectOn]  In   デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するBPMUを0A印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ClampVでクランプ電圧を設定｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■MVRangeで測定レンジを指定。-1(デフォルト)の場合、ClampVからレンジを判定して設定する。
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
    Call SetFIMVMulti(PinList, 0 * A, ClampV, , MVRange, ConnectOn)
End Sub
'#Pass-Ok
Private Sub DisconnectPins(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)
'内容:
'   BPMUをデバイスから切り離す｡
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'    [Site]      In   サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するBPMUをデバイスから切り離す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    
    'Error Check ************************************
    Const FunctionName = "DisconnectPins_BPMU"
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
            '-------------------------------------------------
            
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

End Sub
'#V21-Release
Private Sub DisconnectPinsMulti(ByVal PinList As String)
'内容:
'   BPMUをデバイスから切り離す｡(Site同時)
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するBPMUをデバイスから切り離す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long
    
    'Error Check ************************************
    Const FunctionName = "DisconnectPinsMulti_BPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
        
    'Main Part ---------------------------------------
     TheHdw.Digital.relays.Pins(PinList).DisconnectPins
    '-------------------------------------------------


End Sub
'#Pass-Ok
Private Sub ConnectPins(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)
'内容:
'    BPMUをデバイスに接続する｡
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'    [Site]      In   サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するBPMUをデバイスに接続する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    
    'Error Check ************************************
    Const FunctionName = "DisconnectPins_BPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            TheHdw.BPMU.Pins(PinList).Connect
            '-------------------------------------------------
            
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

End Sub
'#V21-Release
Private Sub ConnectPinsMultiPins(ByVal PinList As String)
'内容:
'    BPMUをデバイスに接続する｡(サイト同時)
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するBPMUをデバイスに接続する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long
    
    'Error Check ************************************
    Const FunctionName = "ConnectPinsMulti_BPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
 
        
            'Main Part ---------------------------------------
            TheHdw.BPMU.Pins(PinList).Connect
            '-------------------------------------------------
            


End Sub
'#V21-Release
Private Sub ConnectPinsMultiChans(ByVal PinList As String)
'内容:
'    BPMUをデバイスに接続する｡(サイト同時)
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するBPMUをデバイスに接続する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long
    Dim Channels() As Long
    
    'Error Check ************************************
    Const FunctionName = "ConnectPinsMulti_BPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
 
        
            'Main Part ---------------------------------------
            
            Call GetActiveChanList(PinList, chIO, Channels)
            TheHdw.BPMU.Chans(Channels).Connect
            '-------------------------------------------------
            


End Sub
'#Pass-Ok
Private Sub SetGND( _
    ByVal PinList As String, Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'   BPMUを0V印加状態に設定する｡
'
'パラメータ:
'    [PinList]     In   対象ピンリスト。
'    [Site]        In   サイト指定。オプション(Default -1)
'    [ConnectOn]   In   デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するBPMUを0V印加状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
    Call SetFVMI(PinList, 0 * V, 200 * mA, site, , , ConnectOn)

End Sub

'#Pass-Ok
Private Sub ChangeMIRange( _
    ByVal PinList As String, ByVal ClampI As Double, Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal MIRange As BpmuIRange = DEF_RANGE _
)
'内容:
'   電圧印加状態のBPMUの電流測定レンジを変更する｡
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'    [ClampI]    In   クランプ電流値。
'    [Site]      In   サイト指定。オプション(Default -1)
'    [MIRange]   In   電流測定レンジ。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するBPMUの印加電圧を変えず､電流測定レンジを変更する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ClampIでクランプ電流を設定｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■MIRangeで測定レンジを指定。-1(デフォルト)の場合、ClampIからレンジを判定して設定する。
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long
    Dim ForceV As Double
    Dim CurIRanges() As Long
    Dim CurVRanges() As Long
    
    Dim ErrPins As String
    
    'Error Check ************************************
    Const FunctionName = "ChangeMIRange_BPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 0, 200 * mA, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    If MIRange = DEF_RANGE Then
        MIRange = GetBpmuMIRange(ClampI)
    End If
    
    'Site Loop **************************************
    ErrPins = ""
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            If TheHdw.BPMU.Pins(PinList).BpmuIsMeasuringCurrent(curSite) = False Then
                ErrPins = ErrPins & PinList & "(site" & curSite & "),"
       '         Call OutputErrMsg(PinList & " is not MI Mode. (at ChangeMIRange_BPMU)")
            Else
                Call GetChanList(PinList, curSite, chIO, Channels)
                For chan = 0 To UBound(Channels)
                    With TheHdw.BPMU.Chans(Channels(chan))
                        Call .ReadDriverRanges(CurIRanges, CurVRanges)
                        .ClampCurrent(MIRange) = ClampI
                        Call .ModeFVMI(MIRange, CurVRanges(0))
                    End With
                Next chan
            End If
            '-------------------------------------------------
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

    If ErrPins <> "" Then
        Call OutputErrMsg(ErrPins & " is not MI Mode. (at ChangeMIRange_BPMU)")
    End If

End Sub
'#V21-Release
Private Sub ChangeMIRangeMulti( _
    ByVal PinList As String, ByVal ClampI As Double, _
    Optional ByVal MIRange As BpmuIRange = DEF_RANGE _
)

'内容:
'   電圧印加状態のBPMUの電流測定レンジを変更する｡ (Site同時)
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'    [ClampI]    In   クランプ電流値。
'    [MIRange]   In   電流測定レンジ。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するBPMUの印加電圧を変えず､電流測定レンジを変更する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ClampIでクランプ電流を設定｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■MIRangeで測定レンジを指定。-1(デフォルト)の場合、ClampIからレンジを判定して設定する。
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long
    Dim Channels() As Long
'    Dim chan As Long
    Dim ForceV As Double
    Dim CurIRanges() As Long
    Dim CurVRanges() As Long
    
    Dim ErrPins As String
    
    'Error Check ************************************
    Const FunctionName = "ChangeMIRangeMulti_BPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 0, 200 * mA, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
    If MIRange = DEF_RANGE Then
        MIRange = GetBpmuMIRange(ClampI)
    End If
    
'    ErrPins = ""

            'MainのエラーCheckは後で再考
            'Main Part ---------------------------------------
'            If TheHdw.BPMU.Pins(pinList).BpmuIsMeasuringCurrent(CurSite) = False Then
'                ErrPins = ErrPins & pinList & "(site" & CurSite & "),"
'       '         Call OutputErrMsg(PinList & " is not MI Mode. (at ChangeMIRange_BPMU)")
'            Else
 'Pins指定
                    With TheHdw.BPMU.Pins(PinList)
                        Call .ReadDriverRanges(CurIRanges, CurVRanges)
                        .ClampCurrent(MIRange) = ClampI
                        Call .ModeFVMI(MIRange, CurVRanges(0))
                    End With
            
 'Chans指定
'                Call GetActiveChanList(pinList, chIO, Channels)
'                    With TheHdw.BPMU.Chans(Channels)
'                        Call .ReadDriverRanges(CurIRanges, CurVRanges)
'                        .ClampCurrent(MIRange) = ClampI
'                        Call .ModeFVMI(MIRange, CurVRanges(0))
'                    End With

'            End If
            '-------------------------------------------------
      
'    If ErrPins <> "" Then
'        Call OutputErrMsg(ErrPins & " is not MI Mode. (at ChangeMIRange_BPMU)")
'    End If


End Sub
'#Pass-Ok
Private Sub MeasureV( _
    ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE _
)
'内容:
'   BPMUの電圧値を測定する｡
'
'パラメータ:
'    [PinName]       In   対象ピン名。
'    [retResult()]   Out  結果を格納する配列。(サイト数分の要素が必要)
'    [AvgNum]        In   平均回数。
'    [Site]          In   サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinNameに対応するBPMUの電圧値を測定する｡
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
    
    Dim ErrPins As String
    
    'Error Check ************************************
    Const FunctionName = "MeasureV_BPMU"
    If CheckPinList(PinName, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckSinglePins(PinName, chIO, FunctionName) = False Then Stop: Exit Sub
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
    ErrPins = ""
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            Call GetChanList(PinName, curSite, chIO, Channels)
            With TheHdw.BPMU
                If .Pins(PinName).BpmuIsMeasuringVoltage(curSite) = False Then
                    ErrPins = ErrPins & PinName & "(site" & curSite & "),"
'                    Call OutputErrMsg(PinName & " is not MV Mode (at MeasureV_BPMU)")
                Else
                    Call .Chans(Channels).measure(avgNum, Samples)
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

    If ErrPins <> "" Then
        Call OutputErrMsg(ErrPins & " is not MV Mode (at MeasureV_BPMU)")
    End If

End Sub

'#Pass-Ok
Private Sub MeasureI( _
    ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE _
)
'内容:
'    BPMUの電流値を測定する｡
'
'パラメータ:
'    [PinName]       In   対象ピン名。
'    [retResult()]   Out  結果を格納する配列。(サイト数分の要素が必要)
'    [AvgNum]        In   平均回数。
'    [Site]          In   サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinNameに対応するBPMUの電流値を測定する｡
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
    
    Dim ErrPins As String
    
    'Error Check ************************************
    Const FunctionName = "MeasureI_BPMU"
    If CheckPinList(PinName, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckSinglePins(PinName, chIO, FunctionName) = False Then Stop: Exit Sub
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
    ErrPins = ""
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            Call GetChanList(PinName, curSite, chIO, Channels)
            With TheHdw.BPMU
                If .Pins(PinName).BpmuIsMeasuringCurrent(curSite) = False Then
                    ErrPins = ErrPins & PinName & "(site" & curSite & "),"
'                    Call OutputErrMsg(PinName & " is not MV Mode (at MeasureI_BPMU)")
                Else
                    Call .Chans(Channels).measure(avgNum, Samples)
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
    
    If ErrPins <> "" Then
        Call OutputErrMsg(ErrPins & " is not MV Mode (at MeasureI_BPMU)")
    End If

End Sub

'#Pass-Ok
Private Function GetBpmuFVRange(ByVal ForceV As Double) As BpmuVRange

    If Abs(ForceV) <= 2 Then        '2V
        GetBpmuFVRange = bpmu2V
    ElseIf Abs(ForceV) <= 5 Then    '5V
        GetBpmuFVRange = bpmu5V
    ElseIf Abs(ForceV) <= 10 Then   '10V
        GetBpmuFVRange = bpmu10V
    Else
        GetBpmuFVRange = bpmu24V
    End If

End Function

'#Pass-Ok
Private Function GetBpmuMIRange(ByVal ClampI As Double) As BpmuIRange

    Dim AbsClampI As Double

    'bpmu2uAレンジは電流クランプ機能が無効になるため、
    'ClampIが2uA以下でもbpmu20uAを設定値として返す。
    
    AbsClampI = Abs(ClampI)
    If AbsClampI <= 0.00002 Then        '20uA
        GetBpmuMIRange = bpmu20uA
    ElseIf AbsClampI <= 0.0002 Then     '200uA
        GetBpmuMIRange = bpmu200uA
    ElseIf AbsClampI <= 0.002 Then      '2mA
        GetBpmuMIRange = bpmu2mA
    ElseIf AbsClampI <= 0.02 Then       '20mA
        GetBpmuMIRange = bpmu20mA
    Else
        GetBpmuMIRange = bpmu200mA
    End If

End Function

'#Pass-Ok
Private Function GetBpmuFIRange(ByVal ForceI As Double) As BpmuIRange

    If Abs(ForceI) <= 0.0002 Then       '200uA
        GetBpmuFIRange = bpmu200uA
    ElseIf Abs(ForceI) <= 0.002 Then    '2mA
        GetBpmuFIRange = bpmu2mA
    ElseIf Abs(ForceI) <= 0.02 Then     '20mA
        GetBpmuFIRange = bpmu20mA
    Else
        GetBpmuFIRange = bpmu200mA
    End If

End Function

'#Pass-Ok
Private Function GetBpmuMVRange(ByVal ClampV As Double) As BpmuVRange

    Dim AbsClampV As Double

    AbsClampV = Abs(ClampV)
    If AbsClampV <= 2 Then          '2V
        GetBpmuMVRange = bpmu2V
    ElseIf AbsClampV <= 5 Then      '5V
        GetBpmuMVRange = bpmu5V
    ElseIf AbsClampV <= 10 Then     '10V
        GetBpmuMVRange = bpmu10V
    Else
        GetBpmuMVRange = bpmu24V
    End If

End Function

'#Pass-OK
Private Function GetBpmuMeasureRange() As Collection
'BPMUの測定レンジを取得する

    Dim pinNameArr() As String
    Dim chListArr() As Long
    Dim tmpBpmuVrange() As Long
    Dim tmpBpmuIrange() As Long
    Dim chkMvMode As Boolean
    Dim pinLpCnt As Long
    Dim siteLpCnt As Long
    Dim tmpMeasureRng() As Double
    Dim tmpRngVal As Collection
                
    Const FUNCTION_NAME = "GetBpmuMeasureRange"
                                
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
        
        '現在のモードと設定されているレンジ値をTERADYNE-APIより取得
        TheHdw.BPMU.Chans(chListArr).ReadDriverRanges tmpBpmuIrange, tmpBpmuVrange
        
        'サイトループ(Chループ)
        For siteLpCnt = 0 To CountExistSite
            With TheHdw.BPMU.Pins(pinNameArr(pinLpCnt))
                '電圧測定 or 電流測定のモード確認
                chkMvMode = .BpmuIsMeasuringVoltage(siteLpCnt)
            End With
            'モード別測定レンジ値選択処理
            If chkMvMode = True Then
                tmpMeasureRng(siteLpCnt) = BpmuVRangeToValue(tmpBpmuVrange(siteLpCnt)) 'MV
            Else
                tmpMeasureRng(siteLpCnt) = BpmuIRangeToValue(tmpBpmuIrange(siteLpCnt)) 'MI
            End If
        Next siteLpCnt
        
        'ピン名をキーにレンジ値の取得結果をコレクションに登録
        Call tmpRngVal.Add(tmpMeasureRng, pinNameArr(pinLpCnt))
    
    Next pinLpCnt
    
    '取得結果を関数の戻り値としてコレクションで返す
    Set GetBpmuMeasureRange = tmpRngVal
        
End Function

'#Pass-Ok
Private Function GetBpmuForceValue() As Collection
'BPMUのForce値を取得する

    Dim pinNameArr() As String
    Dim chListArr() As Long
    Dim tmpBpmuVrange() As Long
    Dim tmpBpmuIrange() As Long
    Dim chkFvMode As Boolean
    Dim pinLpCnt As Long
    Dim siteLpCnt As Long
    Dim tmpForceVal() As Double
    Dim tmpVrange As BpmuVRange
    Dim tmpIrange As BpmuIRange
    Dim tmpForce As Collection
                                                    
    '取得レンジ格納用変数
    ReDim tmpForceVal(CountExistSite)
    
    '最終取得結果用コレクションを準備
    Set tmpForce = New Collection
                
    'カンマ区切りのピンリストを配列形式に変換
    Call SeparatePinList(m_PinList, pinNameArr)
    
    'ピンLOOP
    For pinLpCnt = 0 To UBound(pinNameArr)
        
        'Pinリストからマルチサイト分のチャンネルリストを作成
        Call GetChanList(pinNameArr(pinLpCnt), ALL_SITE, GetChanType(pinNameArr(pinLpCnt)), chListArr)
        
        '現在のモードと設定されているレンジ値をTERADYNE-APIより取得
        TheHdw.BPMU.Chans(chListArr).ReadDriverRanges tmpBpmuIrange, tmpBpmuVrange
        
        'サイトループ(Chループ)
        For siteLpCnt = 0 To CountExistSite
            '電圧印加 or 電流印加のモード確認
            chkFvMode = TheHdw.BPMU.Pins(pinNameArr(pinLpCnt)).BpmuIsForcingVoltage(siteLpCnt)
            With TheHdw.BPMU.Chans(chListArr(siteLpCnt))
                'モード別測定レンジ値選択処理
                If chkFvMode = True Then
                    tmpForceVal(siteLpCnt) = .ForceVoltage(tmpVrange)  'FV
                Else
                    tmpForceVal(siteLpCnt) = .ForceCurrent(tmpIrange)  'FI
                End If
            End With
        Next siteLpCnt
        
        'ピン名をキーにレンジ値の取得結果をコレクションに登録
        Call tmpForce.Add(tmpForceVal, pinNameArr(pinLpCnt))
        
    Next pinLpCnt
    
    '取得結果を関数の戻り値としてコレクションで返す
    Set GetBpmuForceValue = tmpForce
        
End Function

'#Pass-Ok
Private Function BpmuIRangeToValue(ByVal Irange As BpmuIRange) As Double
'BPMUのIRangeをDouble値に変換

    Select Case Irange
        Case bpmuSmartIRange
            BpmuIRangeToValue = DC_RANGE_SMART
        
        Case bpmuAutoIRange
            BpmuIRangeToValue = DC_RANGE_AUTO
        
        Case Else
            BpmuIRangeToValue = TheHdw.BPMU.CurrentRangeToValue(Irange)
    End Select

End Function

'#Pass-OK
Private Function BpmuVRangeToValue(ByVal Vrange As BpmuVRange) As Double
'BPMUのVRangeをDouble値に変換

    Select Case Vrange
        Case bpmuSmartVRange
            BpmuVRangeToValue = DC_RANGE_SMART
        
        Case bpmuAutoVRange
            BpmuVRangeToValue = DC_RANGE_AUTO
        
        Case Else
            BpmuVRangeToValue = TheHdw.BPMU.VoltageRangeToValue(Vrange)
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

    'IG-XLシミュレータで実行された場合はエラーとしNothingを返す
    If TheExec.TesterMode = testModeOffline Then
        Call MsgBox("It is not possible to execute it because there seems to be a problem in API of TERADYNE.", vbCritical, "GetForceValue")
        Exit Function
    Else
        Set IVISVISource_GetForceValue = GetBpmuForceValue
    End If

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
    
    'IG-XLシミュレータで実行された場合はエラーとしNothingを返す
    If TheExec.TesterMode = testModeOffline Then
        Call MsgBox("It is not possible to execute it because there seems to be a problem in API of TERADYNE.", vbCritical, "GetMeasureRange")
        Exit Function
    Else
        Set IVISVISource_GetMeasureRange = GetBpmuMeasureRange
    End If
    
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
    Call SetFVMI(m_PinList, ForceV, Abs(ClampI), site, , , ConnectOn)
End Sub
'V21-Release
Private Sub IVISVISource_SetFVMIMulti(ByVal ForceV As Variant, ByVal ClampI As Double, Optional ByVal ConnectOn As Boolean = True)
    Call SetFVMIMulti(m_PinList, ForceV, Abs(ClampI), , , ConnectOn)
End Sub
Private Sub IVISVISource_SetFIMV(ByVal ForceI As Variant, ByVal ClampV As Double, Optional ByVal site As Long = -1&, Optional ByVal ConnectOn As Boolean = True)
    Call SetFIMV(m_PinList, ForceI, Abs(ClampV), site, , , ConnectOn)
End Sub
'V21-Release
Private Sub IVISVISource_SetFIMVMulti(ByVal ForceI As Variant, ByVal ClampV As Double, Optional ByVal ConnectOn As Boolean = True)
    Call SetFIMVMulti(m_PinList, ForceI, Abs(ClampV), , , ConnectOn)
End Sub
Private Sub IVISVISource_SetMV(ByVal ClampV As Double, Optional ByVal site As Long = -1&, Optional ByVal ConnectOn As Boolean = True)
    Call SetMV(m_PinList, ClampV, site, , ConnectOn)
End Sub
'V21-Release
Private Sub IVISVISource_SetMVMulti(ByVal ClampV As Double, Optional ByVal ConnectOn As Boolean = True)
    Call SetMVMulti(m_PinList, ClampV, , ConnectOn)
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
'#V21-Release
Private Sub IVISVISource_DisconnectPinsMulti()
    
    Call DisconnectPinsMulti(m_PinList)

End Sub
Private Sub IVISVISource_ChangeMIRange(ByVal ClampI As Double, Optional ByVal site As Long = -1&)
    Call ChangeMIRange(m_PinList, Abs(ClampI), site)
End Sub
'#V21-Release
Private Sub IVISVISource_ChangeMIRangeMulti(ByVal ClampI As Double)

    Call ChangeMIRangeMulti(m_PinList, Abs(ClampI))

End Sub
Private Sub IVISVISource_MeasureV(retResult() As Double, ByVal avgNum As Long, Optional ByVal site As Long = -1&)
    Call MeasureV(m_PinList, retResult, avgNum, site)
End Sub

Private Sub IVISVISource_MeasureI(retResult() As Double, ByVal avgNum As Long, Optional ByVal site As Long = -1&)
    Call MeasureI(m_PinList, retResult, avgNum, site)
End Sub

'#################################################################
'BPMUリソースはマルチピンの測定をサポートしない
'（複数ピンの測定を同時に行った際に、それぞれのPINの測定データが
'取得できるかどうかは、ハードウェアの仕様で決まってしまうため）
Private Sub IVISVISource_MeasureIMulti(ByVal avgNum As Long)
    Call OutputErrMsg("MeasureIMulti is not support (at CVISBPMU)")
End Sub

Private Sub IVISVISource_MeasureVMulti(ByVal avgNum As Long)
    Call OutputErrMsg("MeasureVMulti is not support (at CVISBPMU)")
End Sub

Private Sub IVISVISource_ReadMIMultiResult(retResult() As Double)
    Call OutputErrMsg("ReadMIMultiResult is not support (at CVISBPMU)")
End Sub

Private Sub IVISVISource_ReadMVMultiResult(retResult() As Double)
    Call OutputErrMsg("ReadMVMultiResult is not support (at CVISBPMU)")
End Sub
'#################################################################


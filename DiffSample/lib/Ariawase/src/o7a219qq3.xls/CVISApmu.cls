VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVISApmu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   APMUボード制御用クラス
'
'目的:
'   Ⅰ:APMUボード使い勝手向上
'   Ⅱ:V/ISourceとしてのI/F提供
'
'作成者:
'   SLSI今手
'
'XLibSTD_CommonDCModよりAPMU関連機能を整理
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
Private Const BOARD_NAME = "dcAPMU"

Private Const FV_LIMIT_HIGH = 35# * V
Private Const FV_LIMIT_LOW = -35# * V
Private Const FI_LIMIT_HIGH = 50# * mA
Private Const FI_LIMIT_LOW = -50# * mA

Private Const FV_RANGE_LIST = "2V,5V,10V,35V" '装置によっては40Vが設定可能
Private Const MV_RANGE_LIST = "2V,5V,10V,35V" '装置によっては40Vが設定可能
Private Const FI_RANGE_LIST = "40uA,200uA,1mA,5mA,50mA"
Private Const MI_RANGE_LIST = "200nA,2uA,10uA,40uA,200uA,1mA,5mA,50mA"
Private Const GANG_FI_RANGE_LIST = "50mA" 'GANG設定時には電流レンジは50mA固定となる
Private Const GANG_MI_RANGE_LIST = "50mA" 'GANG設定時には電流レンジは50mA固定となる

Private m_PinList As String
Private m_ResultsV_APMU As Collection
Private m_ResultsI_APMU As Collection

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

'#Pass-OK
Private Sub SetFVMI( _
    ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, _
    Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal FVRange As ApmuVRange = DEF_RANGE, Optional ByVal MIRange As ApmuIRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'    APMUを電圧印加状態に設定する｡
'
'パラメータ:
'    [PinList]   In  対象ピンリスト。
'    [ForceV]    In  印加電圧。配列指定可能。
'    [ClampI]    In  クランプ電流値。
'    [Site]      In  サイト指定。オプション(Default -1)
'    [FVRange]   In  電圧印加レンジ。オプション(Default -1)
'    [MIRange]   In  電流測定レンジ。オプション(Default -1)
'    [ConnectOn] In  デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するAPMUを電圧印加状態にする｡
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
    Dim FVRangeBySite() As ApmuVRange

    'Error Check ************************************
    Const FunctionName = "SetFVMI_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceV, -35 * V, 35 * V, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 0 * mA, 50 * mA * 8, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    Call ConvertVariableToArray(ForceV)

    ReDim FVRangeBySite(CountExistSite)
    For curSite = 0 To CountExistSite
        If FVRange = DEF_RANGE Then
            FVRangeBySite(curSite) = GetApmuFVRange(ForceV(curSite))
        Else
            FVRangeBySite(curSite) = FVRange
        End If
    Next curSite
    
    If MIRange = DEF_RANGE Then
        MIRange = GetApmuMIRange(ClampI)
    End If
    
    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            With TheHdw.APMU.Pins(PinList)
                .alarm = False
                .ClampCurrent(MIRange) = ClampI
                .ForceVoltage(FVRangeBySite(curSite)) = ForceV(curSite)
                Call .ModeFVMI(MIRange)
                .Gate = True
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
    Optional ByVal FVRange As ApmuVRange = DEF_RANGE, Optional ByVal MIRange As ApmuIRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
    )
'内容:
'    APMUを電圧印加状態に設定する｡(Site同時)
'
'パラメータ:
'    [PinList]   In  対象ピンリスト。
'    [ForceV]    In  印加電圧。配列指定可能。
'    [ClampI]    In  クランプ電流値。
'    [FVRange]   In  電圧印加レンジ。オプション(Default -1)
'    [MIRange]   In  電流測定レンジ。オプション(Default -1)
'    [ConnectOn] In  デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するAPMUを電圧印加状態にする｡
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
    Dim FVRangeBySite() As ApmuVRange
    Dim Channels() As Long

    'Error Check ************************************
    Const FunctionName = "SetFVMIMulti_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceV, -35 * V, 35 * V, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 0 * mA, 50 * mA * 8, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
'site0only設定************************************
    Call ConvertVariableToArray(ForceV)
    ReDim FVRangeBySite(1)
    curSite = 0
        If FVRange = DEF_RANGE Then
            FVRangeBySite(curSite) = GetApmuFVRange(ForceV(curSite))
        Else
            FVRangeBySite(curSite) = FVRange
        End If
  '************************************************
    
    If MIRange = DEF_RANGE Then
        MIRange = GetApmuMIRange(ClampI)
    End If
    
    
        
            'Main Part ---------------------------------------
' Pins指定
'            With TheHdw.APMU.Pins(pinList)
'                .Alarm = False
'                .ClampCurrent(MIRange) = ClampI
'                .ForceVoltage(FVRangeBySite(CurSite)) = ForceV(CurSite)
'                Call .ModeFVMI(MIRange)
'                .Gate = True
'            End With
            
'Chans指定
            Call GetActiveChanList(PinList, chAPMU, Channels)
            With TheHdw.APMU.Chans(Channels)
                .alarm = False
                .ClampCurrent(MIRange) = ClampI
                .ForceVoltage(FVRangeBySite(curSite)) = ForceV(curSite)
                Call .ModeFVMI(MIRange)
                .Gate = True
            End With
            '-------------------------------------------------
            
        

    If ConnectOn = True Then
        Call ConnectPinsMultiPins(PinList)
'        Call ConnectPinsMultiChans(pinList)
    End If
End Sub
'#Pass-Ok
Private Sub SetFIMV( _
    ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, _
    Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal FIRange As ApmuIRange = DEF_RANGE, Optional ByVal MVRange As ApmuVRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'    APMUを電流印加状態に設定する｡
'
'パラメータ:
'    [PinList]   In  対象ピンリスト。
'    [ForceI]    In  印加電流。配列指定可能。
'    [ClampV]    In  クランプ電圧値。
'    [Site]      In  サイト指定。オプション(Default -1)
'    [FIRange]   In  電流印加レンジ。オプション(Default -1)
'    [MVRange]   In  電圧測定レンジ。オプション(Default -1)
'    [ConnectOn] In  デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するAPMUを電流印加状態にする｡
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
    Dim FIRangeBySite() As ApmuIRange

    'Error Check ************************************
    Const FunctionName = "SetFIMV_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceI, -50 * mA * 8, 50 * mA * 8, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampV, 0 * V, 35 * V, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    Call ConvertVariableToArray(ForceI)
    
    ReDim FIRangeBySite(CountExistSite)
    For curSite = 0 To CountExistSite
        If FIRange = DEF_RANGE Then
            FIRangeBySite(curSite) = GetApmuFIRange(ForceI(curSite))
        Else
            FIRangeBySite(curSite) = FIRange
        End If
    Next curSite
    
    If MVRange = DEF_RANGE Then
        MVRange = GetApmuMVRange(ClampV)
    End If
    
    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            With TheHdw.APMU.Pins(PinList)
                .alarm = False
                .ClampVoltage(MVRange) = ClampV
                .ForceCurrent(FIRangeBySite(curSite)) = ForceI(curSite)
                Call .ModeFIMV(MVRange)
                .Gate = True
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
    Optional ByVal FIRange As ApmuIRange = DEF_RANGE, Optional ByVal MVRange As ApmuVRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)

'内容:
'    APMUを電流印加状態に設定する｡(Site同時)
'
'パラメータ:
'    [PinList]   In  対象ピンリスト。
'    [ForceI]    In  印加電流。配列指定可能。
'    [ClampV]    In  クランプ電圧値。
'    [FIRange]   In  電流印加レンジ。オプション(Default -1)
'    [MVRange]   In  電圧測定レンジ。オプション(Default -1)
'    [ConnectOn] In  デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するAPMUを電流印加状態にする｡
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
    Dim FIRangeBySite() As ApmuIRange
    Dim Channels() As Long

    'Error Check ************************************
    Const FunctionName = "SetFIMVMulti_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceI, -50 * mA * 8, 50 * mA * 8, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampV, 0 * V, 35 * V, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
   'site0only設定************************************
    Call ConvertVariableToArray(ForceI)
    ReDim FIRangeBySite(1)
        curSite = 0
        If FIRange = DEF_RANGE Then
            FIRangeBySite(curSite) = GetApmuFIRange(ForceI(curSite))
        Else
            FIRangeBySite(curSite) = FIRange
        End If
   '************************************************
    
    If MVRange = DEF_RANGE Then
        MVRange = GetApmuMVRange(ClampV)
    End If
    

        
            'Main Part ---------------------------------------
'Pins指定
'            With TheHdw.APMU.Pins(pinList)
'                .Alarm = False
'                .ClampVoltage(MVRange) = ClampV
'                .ForceCurrent(FIRangeBySite(CurSite)) = ForceI(CurSite)
'                Call .ModeFIMV(MVRange)
'                .Gate = True
'            End With
            
'Chans指定
            Call GetActiveChanList(PinList, chAPMU, Channels)
            With TheHdw.APMU.Chans(Channels)
                .alarm = False
                .ClampVoltage(MVRange) = ClampV
                .ForceCurrent(FIRangeBySite(curSite)) = ForceI(curSite)
                Call .ModeFIMV(MVRange)
                .Gate = True
            End With
            '-------------------------------------------------


    If ConnectOn = True Then
        Call ConnectPinsMultiPins(PinList)
        'Call ConnectPinsMultiChans(pinList)
    End If


End Sub
'#Pass-Ok
Private Sub SetMV( _
    ByVal PinList As String, ByVal ClampV As Double, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal MVRange As ApmuVRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'    APMUを電圧測定状態(無負荷)に設定する。
'
'パラメータ:
'    [PinList]    In  対象ピンリスト。
'    [ClampV]     In  クランプ電圧値。
'    [Site]       In  サイト指定。オプション(Default -1)
'    [MVRange]    In  電圧測定レンジ。オプション(Default -1)
'    [ConnectOn]  In  デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するAPMUを無負荷電圧測定状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ClampVでクランプ電圧を設定｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■MVRangeで測定レンジを指定。-1(デフォルト)の場合、ClampVからレンジを判定して設定する。
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long

    
    'Error Check ************************************
    Const FunctionName = "SetMV_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampV, 0 * V, 35 * V, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    If MVRange = DEF_RANGE Then
        MVRange = GetApmuMVRange(ClampV)
    End If
    
    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            With TheHdw.APMU.Pins(PinList)
                .alarm = False
                .ClampVoltage(MVRange) = ClampV
                Call .ModeMV(MVRange)
                .Gate = True
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
Private Sub SetMVMulti(ByVal PinList As String, ByVal ClampV As Double, _
    Optional ByVal MVRange As ApmuVRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)

'内容:
'    APMUを電圧測定状態(無負荷)に設定する。(Site同時)
'
'パラメータ:
'    [PinList]    In  対象ピンリスト。
'    [ClampV]     In  クランプ電圧値。
'    [MVRange]    In  電圧測定レンジ。オプション(Default -1)
'    [ConnectOn]  In  デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するAPMUを無負荷電圧測定状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ClampVでクランプ電圧を設定｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■MVRangeで測定レンジを指定。-1(デフォルト)の場合、ClampVからレンジを判定して設定する。
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
'    Dim SiteStatus As loopStatus
    Dim curSite As Long
    Dim Channels() As Long
    
    'Error Check ************************************
    Const FunctionName = "SetMVMulti_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampV, 0 * V, 35 * V, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
    If MVRange = DEF_RANGE Then
        MVRange = GetApmuMVRange(ClampV)
    End If
 
        
            'Main Part ---------------------------------------
' Pins指定
            With TheHdw.APMU.Pins(PinList)
                .alarm = False
                .ClampVoltage(MVRange) = ClampV
                Call .ModeMV(MVRange)
                .Gate = True
            End With
'Chans指定
'            Call GetActiveChanList(pinList, chAPMU, Channels)
'            With TheHdw.APMU.Chans(Channels)
'             .Alarm = False
'                .ClampVoltage(MVRange) = ClampV
'                Call .ModeMV(MVRange)
'                .Gate = True
'            End With
            '-------------------------------------------------


    If ConnectOn = True Then
        Call ConnectPinsMultiPins(PinList)
        'Call ConnectPinsMultiChans(pinList)
    End If


End Sub
'#Pass-OK
Private Sub DisconnectPins(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)
'内容:
'    APMUをデバイスから切り離す｡
'
'パラメータ:
'    [PinList]  In  対象ピンリスト。
'    [Site]     In  サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するAPMUをデバイスから切り離す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■切り離した後、0V印加状態に設定する(接続はしない)。
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    
    'Error Check ************************************
    Const FunctionName = "DisconnectPins_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            With TheHdw.APMU.Pins(PinList)
                .relay = False
                .Gate = False
                Call .ModeFVMI(apmu50mA)
                .ForceVoltage(apmu2V) = 0 * V
                .ClampCurrent(apmu50mA) = 50 * mA
                .alarm = True
            End With
            '-------------------------------------------------
            
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

End Sub
'#V21-Relese
Private Sub DisconnectPinsMulti(ByVal PinList As String)
'内容:
'    APMUをデバイスから切り離す｡(Site同時)
'
'パラメータ:
'    [PinList]  In  対象ピンリスト。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するAPMUをデバイスから切り離す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■切り離した後、0V印加状態に設定する(接続はしない)。
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long
    Dim Channels() As Long
    
    'Error Check ************************************
    Const FunctionName = "DisconnectPinsMulti_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
        
            'Main Part ---------------------------------------
'Pins指定
'            With TheHdw.APMU.Pins(pinList)
'                .relay = False
'                .Gate = False
'                Call .ModeFVMI(apmu50mA)
'                .ForceVoltage(apmu2V) = 0 * V
'                .ClampCurrent(apmu50mA) = 50 * mA
'                .Alarm = True
'            End With
            
'Chans指定
            Call GetActiveChanList(PinList, chAPMU, Channels)
            With TheHdw.APMU.Chans(Channels)
                .relay = False
                .Gate = False
                Call .ModeFVMI(apmu50mA)
                .ForceVoltage(apmu2V) = 0 * V
                .ClampCurrent(apmu50mA) = 50 * mA
                .alarm = True
            End With
            '-------------------------------------------------
            

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
            Call GetChanList(TmpPinList, curSite, chAPMU, Channels)
            For chan = 0 To UBound(Channels)
                TheHdw.APMU.Chans(Channels(chan)).GangedChannels = False
            Next
        End If
    Next curSite
    
End Sub

'#Pass-OK
Private Sub ConnectPins(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)
'内容:
'   APMUをデバイスに接続する｡
'
'パラメータ:
'    [PinList]  In  対象ピンリスト。
'    [Site]     In  サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するAPMUをデバイスに接続する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    
    'Error Check ************************************
    Const FunctionName = "ConnectPins_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            TheHdw.APMU.Pins(PinList).relay = True
            '-------------------------------------------------
            
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

End Sub
'#V21-Release
Private Sub ConnectPinsMultiPins(ByVal PinList As String)
'内容:
'   APMUをデバイスに接続する｡(サイト同時)
'
'パラメータ:
'    [PinList]  In  対象ピンリスト。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するAPMUをデバイスに接続する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long
    
    'Error Check ************************************
    Const FunctionName = "ConnectPinsMulti_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
        
            'Main Part ---------------------------------------
            TheHdw.APMU.Pins(PinList).relay = True
            '-------------------------------------------------


End Sub
'#V21-Release
Private Sub ConnectPinsMultiChans(ByVal PinList As String)
'内容:
'   APMUをデバイスに接続する｡(サイト同時)
'
'パラメータ:
'    [PinList]  In  対象ピンリスト。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するAPMUをデバイスに接続する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long
     Dim Channels() As Long
    
    'Error Check ************************************
    Const FunctionName = "ConnectPinsMultui_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
        
            'Main Part ---------------------------------------
            Call GetActiveChanList(PinList, chAPMU, Channels)
            TheHdw.APMU.Chans(PinList).relay = True
            '-------------------------------------------------


End Sub


'#Pass-Ok
Private Sub SetGND( _
    ByVal PinList As String, Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'    APMUを接地する｡
'
'パラメータ:
'    [PinList]    In  対象ピンリスト。
'    [Site]       In  サイト指定。オプション(Default -1)
'    [ConnectOn]  In  デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するAPMUを接地する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると、設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
    Call SetFVMI(PinList, 0 * V, 50 * mA, site, , , ConnectOn)
    
End Sub

'#Pass-Ok
Private Sub ChangeMIRange( _
    ByVal PinList As String, ByVal ClampI As Double, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal MIRange As ApmuIRange = DEF_RANGE _
)
'内容:
'   電圧印加状態のAPMUの電流測定レンジを変更する｡
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
'    ■PinListに対応するAPMUの印加電圧を変えず､電流測定レンジを変更する｡
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
    
    Dim CurMode() As ApmuMode
    Dim CurVRange() As ApmuVRange
    Dim CurIRange() As ApmuIRange

    Dim ErrChans As String

    'Error Check ************************************
    Const FunctionName = "ChangeMIRange_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 0 * mA, 50 * mA * 8, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    If MIRange = DEF_RANGE Then
        MIRange = GetApmuMIRange(ClampI)
    End If
    
    'Site Loop **************************************
    ErrChans = ""
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            Call GetChanList(PinList, curSite, chAPMU, Channels)
            For chan = 0 To UBound(Channels)
                With TheHdw.APMU.Chans(Channels(chan))
                    Call .ReadRangesAndMode(CurMode, CurVRange, CurIRange)
                    If CurMode(0) <> apmuForceVMeasureI Then
                        ErrChans = ErrChans & Channels(chan) & ","
'                        Call OutputErrMsg("APMU(" & Channels(chan) & ") is not MI Mode. (at ChangeMIRange_APMU)")
                    Else
                        .ClampCurrent(MIRange) = ClampI
                        Call .ModeFVMI(MIRange)
                    End If
                End With
            Next chan
            '-------------------------------------------------
            
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************
    
    If ErrChans <> "" Then
        Call OutputErrMsg("APMU(" & ErrChans & ") is not MI Mode. (at ChangeMIRange_APMU)")
    End If

End Sub
'#V21-Release
Private Sub ChangeMIRangeMulti( _
    ByVal PinList As String, ByVal ClampI As Double, _
    Optional ByVal MIRange As ApmuIRange = DEF_RANGE _
)

'内容:
'   電圧印加状態のAPMUの電流測定レンジを変更する｡(サイト同時)
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'    [ClampI]    In   クランプ電流値。
'    [MIRange]   In   電流測定レンジ。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するAPMUの印加電圧を変えず､電流測定レンジを変更する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ClampIでクランプ電流を設定｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■MIRangeで測定レンジを指定。-1(デフォルト)の場合、ClampIからレンジを判定して設定する。
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
'    Dim SiteStatus As loopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long
    
    Dim CurMode() As ApmuMode
    Dim CurVRange() As ApmuVRange
    Dim CurIRange() As ApmuIRange

    Dim ErrChans As String

    'Error Check ************************************
    Const FunctionName = "ChangeMIRangeMulti_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 0 * mA, 50 * mA * 8, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
    If MIRange = DEF_RANGE Then
        MIRange = GetApmuMIRange(ClampI)
    End If
    
         
            'Main Part ---------------------------------------
 'Pins指定
'                With TheHdw.APMU.Pins(pinList)
'                    Call .ReadRangesAndMode(CurMode, CurVRange, CurIRange)
'                    If CurMode(0) <> apmuForceVMeasureI Then
'                        ErrChans = ErrChans & Channels(chan) & ","
'                    Else
'                        .ClampCurrent(MIRange) = ClampI
'                        Call .ModeFVMI(MIRange)
'                    End If
'                End With
            
 'Chans指定
            Call GetActiveChanList(PinList, chAPMU, Channels)
                With TheHdw.APMU.Chans(Channels)
                    Call .ReadRangesAndMode(CurMode, CurVRange, CurIRange)
                    If CurMode(0) <> apmuForceVMeasureI Then
                        ErrChans = ErrChans & Channels(chan) & ","
                    Else
                        .ClampCurrent(MIRange) = ClampI
                        Call .ModeFVMI(MIRange)
                    End If
                End With
            
            '-------------------------------------------------
               
    If ErrChans <> "" Then
        Call OutputErrMsg("APMU(" & ErrChans & ") is not MI Mode. (at ChangeMIRange_APMU)")
    End If
End Sub
'#Pass-Ok
Private Sub MeasureV( _
    ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal UseLPF As Boolean = False _
)
'内容:
'    APMUの電圧値を測定する｡
'
'パラメータ:
'    [PinName]      In   対象ピン名。
'    [retResult()]  Out  結果を格納する配列。(サイト数分の要素が必要)
'    [AvgNum]       In   平均回数。
'    [Site]         In   サイト指定。オプション(Default -1)
'    [UseLPF]       In   ローパスフィルタの使用。オプション(Default False)
'
'戻り値:
'
'注意事項:
'    ■PinNameに対応するAPMUの電圧値を測定する｡
'    ■PinNameは１ピンのみ。複数ピン、ピングループ(ギャングピンを除く)の指定は不可。
'    ■retResultに結果を返す｡
'    ■AvgNumの数だけサンプリングして､その平均を返す｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■Siteが-1以外の場合、対象サイト以外のretResult()の値は変化させない。
'    ■非アクティブサイトの結果は0を返す。Siteが-1の時も同様。
'    ■UseLPFをTrueにすると､サンプリング時にローパスフィルタがかかる｡
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim Samples() As Double
    Dim n As Long
    
    Dim CurMode() As ApmuMode
    Dim CurVRange() As ApmuVRange
    Dim CurIRange() As ApmuIRange
    
    Dim ErrPins As String

    'Error Check ************************************
    Const FunctionName = "MeasureV_APMU"
    If CheckPinList(PinName, chAPMU, FunctionName) = False Then Stop: Exit Sub
    If CheckSinglePins(PinName, chAPMU, FunctionName) = False Then Stop: Exit Sub
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
            Call GetChanList(PinName, curSite, chAPMU, Channels)
            
            With TheHdw.APMU.Pins(PinName)
                Call .ReadRangesAndMode(CurMode, CurVRange, CurIRange)
                If CurMode(0) = apmuForceVMeasureI Then
                    ErrPins = ErrPins & PinName & "(site" & curSite & "),"
'                    Call OutputErrMsg(PinName & " is not MV Mode (at MeasureV_APMU)")
                Else
                    .LowPassFilter = UseLPF
                    Call .measure(avgNum, Samples)
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
        Call OutputErrMsg(ErrPins & " is not MV Mode (at MeasureV_APMU)")
    End If

End Sub

'#Pass-Ok
Private Sub MeasureI( _
    ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal UseLPF As Boolean = False _
)
'内容:
'   APMUの電流値を測定する｡
'
'パラメータ:
'    [PinName]       In   対象ピン名。
'    [retResult()]   Out  結果を格納する配列。(サイト数分の要素が必要)
'    [AvgNum]        In   平均回数。
'    [Site]          In   サイト指定。オプション(Default -1)
'    [UseLPF]        In   ローパスフィルタの使用。オプション(Default False)
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
'    ■UseLPFをTrueにすると､サンプリング時にローパスフィルタがかかる｡
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim Samples() As Double
    Dim n As Long
    
    Dim CurMode() As ApmuMode
    Dim CurVRange() As ApmuVRange
    Dim CurIRange() As ApmuIRange

    Dim ErrPins As String
    
    'Error Check ************************************
    Const FunctionName = "MeasureI_APMU"
    If CheckPinList(PinName, chAPMU, FunctionName) = False Then Stop: Exit Sub
    If CheckSinglePins(PinName, chAPMU, FunctionName) = False Then Stop: Exit Sub
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
            Call GetChanList(PinName, curSite, chAPMU, Channels)
             
            With TheHdw.APMU.Pins(PinName)
                Call .ReadRangesAndMode(CurMode, CurVRange, CurIRange)
                If CurMode(0) <> apmuForceVMeasureI Then
                    ErrPins = ErrPins & PinName & "(site" & curSite & "),"
            '                    Call OutputErrMsg(PinName & " is not MI Mode (at MeasureI_APMU)")
                Else
                    .LowPassFilter = UseLPF
                    Call .measure(avgNum, Samples)
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
        Call OutputErrMsg(ErrPins & " is not MI Mode (at MeasureI_APMU)")
    End If

End Sub

'#Pass-Ok
Private Sub MeasureVMulti(ByVal PinList As String, ByVal avgNum As Long, Optional ByVal UseLPF As Boolean = False)
'内容:
'   APMUの電圧値を測定する｡
'
'パラメータ:
'    [PinList]   In   対象ピン名。
'    [AvgNum]    In   平均回数。
'    [UseLPF]    In   ローパスフィルタの使用。オプション(Default False)
'
'戻り値:
'
'詳細動作
'    ■PinListに対応するピンの電圧値を測定する｡結果はReadMVMultiResultを使って読み出す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ギャングピンは対応しない｡
'    ■AvgNumの数だけサンプリングして､その平均を返す｡
'    ■全サイトが対象になる｡非アクティブサイトの結果は0になる｡
'    ■UseLPFをTrueにすると､サンプリング時にローパスフィルタがかかる｡
'
    Dim Channels() As Long
    Dim Samples() As Double
    Dim pinNames() As String
    Dim PerPinResults() As Variant
    Dim curSite As Long
    Dim Pin As Long
    Dim n As Long
    Dim i As Long
    
    Dim CurMode() As ApmuMode
    Dim CurVRange() As ApmuVRange
    Dim CurIRange() As ApmuIRange
    Dim chan As Long
    
    
    'Error Check ************************************
    Const FunctionName = "MeasureVMulti_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
    If CheckAvgNum(avgNum, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    'Check Ganged Pins ********************************************
    Call GetActiveChanList(PinList, chAPMU, Channels)
    Call SeparatePinList(PinList, pinNames)
    If (UBound(pinNames) + 1) * CountActiveSite <> UBound(Channels) + 1 Then
        Call OutputErrMsg(PinList & " Including Ganged Pins. (at MeasureVMulti_APMU)")
        Exit Sub
    End If
    '**************************************************************
    
    'Check MV Mode ************************************************
    Call TheHdw.APMU.Pins(PinList).ReadRangesAndMode(CurMode, CurVRange, CurIRange)
    For chan = 0 To UBound(CurMode)
        If CurMode(chan) = apmuForceVMeasureI Then
            Call OutputErrMsg(PinList & " Including not MV Mode Pins. (at MeasureVMulti_APMU)")
            Exit Sub
        End If
    Next chan
    '**************************************************************
    
    'Measurement **************************************************
    With TheHdw.APMU.Pins(PinList)
        .LowPassFilter = UseLPF
        Call .measure(avgNum, Samples)
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
    Set m_ResultsV_APMU = New Collection
    For Pin = 0 To UBound(pinNames)
        For curSite = 0 To CountExistSite
            PerPinResults(Pin)(curSite) = PerPinResults(Pin)(curSite) / avgNum
        Next curSite
        Call m_ResultsV_APMU.Add(PerPinResults(Pin), pinNames(Pin))
    Next Pin
    '---------------------------------------------------------
    '**************************************************************
    
End Sub

'#Pass-Ok
Private Sub MeasureIMulti(ByVal PinList As String, ByVal avgNum As Long, Optional ByVal UseLPF As Boolean = False)
'内容:
'   APMUの電流値を測定する｡
'
'パラメータ:
'    [PinList]    In   対象ピン名。
'    [AvgNum]     In   平均回数。
'    [UseLPF]     In   ローパスフィルタの使用。オプション(Default False)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するピンの電流値を測定する｡結果はReadMIMultiResultを使って読み出す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■ギャングピンは対応しない｡
'    ■AvgNumの数だけサンプリングして､その平均を返す｡
'    ■全サイトが対象になる｡非アクティブサイトの結果は0になる｡
'    ■UseLPFをTrueにすると､サンプリング時にローパスフィルタがかかる｡
'
    Dim Channels() As Long
    Dim Samples() As Double
    Dim pinNames() As String
    Dim PerPinResults() As Variant
    Dim curSite As Long
    Dim Pin As Long
    Dim n As Long
    Dim i As Long
    
    Dim CurMode() As ApmuMode
    Dim CurVRange() As ApmuVRange
    Dim CurIRange() As ApmuIRange
    Dim chan As Long
    
    
    'Error Check ************************************
    Const FunctionName = "MeasureIMulti_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
    If CheckAvgNum(avgNum, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    'Check Ganged Pins ********************************************
    Call GetActiveChanList(PinList, chAPMU, Channels)
    Call SeparatePinList(PinList, pinNames)
    If (UBound(pinNames) + 1) * CountActiveSite <> UBound(Channels) + 1 Then
        Call OutputErrMsg(PinList & " Including Ganged Pins. (at MeasureIMulti_APMU)")
        Exit Sub
    End If
    '**************************************************************
    
    'Check MI Mode ************************************************
    Call TheHdw.APMU.Pins(PinList).ReadRangesAndMode(CurMode, CurVRange, CurIRange)
    For chan = 0 To UBound(CurMode)
        If CurMode(chan) <> apmuForceVMeasureI Then
            Call OutputErrMsg(PinList & " Including not MI Mode Pins. (at MeasureIMulti_APMU)")
            Exit Sub
        End If
    Next chan
    '**************************************************************
    
    'Measurement **************************************************
    With TheHdw.APMU.Pins(PinList)
        .LowPassFilter = UseLPF
        Call .measure(avgNum, Samples)
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
    Set m_ResultsI_APMU = New Collection
    For Pin = 0 To UBound(pinNames)
        For curSite = 0 To CountExistSite
            PerPinResults(Pin)(curSite) = PerPinResults(Pin)(curSite) / avgNum
        Next curSite
        Call m_ResultsI_APMU.Add(PerPinResults(Pin), pinNames(Pin))
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

    status = ReadMultiResult(PinName, retResult, m_ResultsV_APMU)
        
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
    
    status = ReadMultiResult(PinName, retResult, m_ResultsI_APMU)

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
    Set m_ResultsV_APMU = Nothing
    Set m_ResultsI_APMU = Nothing

End Sub

'#Pass-Ok
Private Function GetApmuFVRange(ByVal ForceV As Double) As ApmuVRange

    If Abs(ForceV) <= 2 Then        '2V
        GetApmuFVRange = apmu2V
    ElseIf Abs(ForceV) <= 5 Then    '5V
        GetApmuFVRange = apmu5V
    ElseIf Abs(ForceV) <= 10 Then   '10V
        GetApmuFVRange = apmu10V
    Else
        GetApmuFVRange = apmu35V
    End If

End Function

'#Pass-Ok
Private Function GetApmuMIRange(ByVal ClampI As Double) As ApmuIRange

    Dim AbsClampI As Double

    AbsClampI = Abs(ClampI)
    If AbsClampI <= 0.0000002 Then      '200nA
        GetApmuMIRange = apmu200nA
    ElseIf AbsClampI <= 0.000002 Then   '2uA
        GetApmuMIRange = apmu2uA
    ElseIf AbsClampI <= 0.00001 Then    '10uA
        GetApmuMIRange = apmu10uA
    ElseIf AbsClampI <= 0.00004 Then    '40uA
        GetApmuMIRange = apmu40uA
    ElseIf AbsClampI <= 0.0002 Then     '200uA
        GetApmuMIRange = apmu200uA
    ElseIf AbsClampI <= 0.001 Then      '1mA
        GetApmuMIRange = apmu1mA
    ElseIf AbsClampI <= 0.005 Then      '5mA
        GetApmuMIRange = apmu5mA
    Else
        GetApmuMIRange = apmu50mA
    End If

End Function

'#Pass-Ok
Private Function GetApmuFIRange(ByVal ForceI As Double) As ApmuIRange

    If Abs(ForceI) <= 0.00004 Then      '40uA
        GetApmuFIRange = apmu40uA
    ElseIf Abs(ForceI) <= 0.0002 Then   '200uA
        GetApmuFIRange = apmu200uA
    ElseIf Abs(ForceI) <= 0.001 Then    '1mA
        GetApmuFIRange = apmu1mA
    ElseIf Abs(ForceI) <= 0.005 Then    '5mA
        GetApmuFIRange = apmu5mA
    Else
        GetApmuFIRange = apmu50mA
    End If

End Function

'#Pass-Ok
Private Function GetApmuMVRange(ByVal ClampV As Double) As ApmuVRange

    Dim AbsClampV As Double

    AbsClampV = Abs(ClampV)
    If AbsClampV <= 2 Then          '2V
        GetApmuMVRange = apmu2V
    ElseIf AbsClampV <= 5 Then      '5V
        GetApmuMVRange = apmu5V
    ElseIf AbsClampV <= 10 Then     '10V
        GetApmuMVRange = apmu10V
    Else
        GetApmuMVRange = apmu35V
    End If

End Function

'#Pass-Ok
Private Function GetApmuMeasureRange() As Collection
'APMUの測定レンジを取得する

    Dim pinNameArr() As String
    Dim chListArr() As Long
    Dim tmpApmuMode() As ApmuMode
    Dim tmpApmuVrange() As ApmuVRange
    Dim tmpApmuIrange() As ApmuIRange
    Dim pinLpCnt As Long
    Dim siteLpCnt As Long
    Dim tmpMeasureRng() As Double
    Dim tmpGangRly() As Long
    Dim tmpRngVal As Collection
    Dim tmpModeVal As Collection
                                                
    Const FUNCTION_NAME = "GetApmuMeasureRange"
                                                                                                    
    '取得レンジ格納用変数
    ReDim tmpMeasureRng(CountExistSite)
    
    '最終取得結果用コレクションを準備
    Set tmpRngVal = New Collection
    
    '印加モード取得結果用コレクションを準備
    Set tmpModeVal = New Collection
                                                
    'カンマ区切りのピンリストを配列形式に変換
    Call SeparatePinList(m_PinList, pinNameArr)
    
    'ピンLOOP
    For pinLpCnt = 0 To UBound(pinNameArr)
        'Pinリストからマルチサイト分のチャンネルリストを作成
        Call GetChanList(pinNameArr(pinLpCnt), ALL_SITE, GetChanType(pinNameArr(pinLpCnt)), chListArr)
        
        '現在のモードと設定されているレンジ値をTERADYNE-APIより取得
        TheHdw.APMU.Chans(chListArr).ReadRangesAndMode tmpApmuMode, tmpApmuVrange, tmpApmuIrange
                
        'サイトループ(CHループ)
        For siteLpCnt = 0 To CountExistSite
            '現在の設定モードに応じて電圧レンジ値もしくは電流レンジ値を格納する
            Select Case tmpApmuMode(siteLpCnt)
                Case apmuForceIMeasureV: 'FIMV
                    tmpMeasureRng(siteLpCnt) = ApmuVRangeToValue(tmpApmuVrange(siteLpCnt))
                
                Case apmuForceVMeasureI: 'FVMI
                    tmpMeasureRng(siteLpCnt) = ApmuIRangeToValue(tmpApmuIrange(siteLpCnt))
                
                Case apmuMeasureV:       'MV
                    tmpMeasureRng(siteLpCnt) = ApmuVRangeToValue(tmpApmuVrange(siteLpCnt))
            End Select
        Next siteLpCnt
        
        'ピン名をキーにレンジ値の取得結果をコレクションに登録
        Call tmpRngVal.Add(tmpMeasureRng, pinNameArr(pinLpCnt))
        'ピン名をキーに印加モードの取得結果をコレクションに登録
        Call tmpModeVal.Add(tmpApmuMode, pinNameArr(pinLpCnt))
    
    Next pinLpCnt
        
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
            Set GetApmuMeasureRange = Nothing
            Exit Function
        End If
                                                                             
        '★ギャング対象GPのメンバーはすべて同じ状態と信じて
        '  先頭PINのレンジ値と印加モードのみを調べる
        ' （本来はすべてのメンバーの状態から判断を行うのが正しい）
        Dim tmpGangMeasureRng() As Double
        Dim tmpGangApmuMode() As ApmuMode
        tmpGangMeasureRng = tmpRngVal.Item(pinNameArr(0))
        tmpGangApmuMode = tmpModeVal.Item(pinNameArr(0))

        '★モード別処理（MIモード時のみ電流レンジ×並列接続数を演算）
        '  (ここでもサイト0の値のみを確認してモードを判断するが
        '  本来はすべてのサイトのモード値を確認するのが正しい
        '  上記のメンバー確認のタイミングですべて確認済みのものを
        '  判定するのが良さそう。)
        Dim tmpGangRngVal As Collection
        Set tmpGangRngVal = New Collection

        If tmpGangApmuMode(0) <> apmuForceVMeasureI Then
            'Gang用のGp名をキーに測定レンジ結果を格納(MVモード時)
            Call tmpGangRngVal.Add(tmpGangMeasureRng, m_PinList)
        Else
            'サイトループ(CHループ)で測定レンジ値×並列接続数を計算（MIモード時）
            For siteLpCnt = 0 To CountExistSite
                tmpGangMeasureRng(siteLpCnt) = tmpGangMeasureRng(siteLpCnt) * (UBound(pinNameArr) + 1)
            Next siteLpCnt
            'Gang用のGp名をキーに測定レンジ結果を格納
            Call tmpGangRngVal.Add(tmpGangMeasureRng, m_PinList)
        End If
                
        'Gang用に計算された測定レンジ値を取得結果として返す
        Set GetApmuMeasureRange = tmpGangRngVal
                
    'ピンリストの中身が通常ピンのとき(取得結果をそのまま返す)
    Else
        '取得結果を関数の戻り値としてコレクションで返す
        Set GetApmuMeasureRange = tmpRngVal

    End If
            
End Function

'#Pass-Ok
Private Function GetApmuForceValue() As Collection
'APMUのForce値を取得する

    Dim pinNameArr() As String
    Dim chListArr() As Long
    Dim tmpApmuMode() As ApmuMode
    Dim tmpApmuVrange() As ApmuVRange
    Dim tmpApmuIrange() As ApmuIRange
    Dim pinLpCnt As Long
    Dim siteLpCnt As Long
    Dim tmpRetForceVal() As Double
    Dim getForceVal() As Double
    Dim tmpForceVal As Collection
    Dim status As Boolean
                                                                                                
    Const FUNCTION_NAME = "GetApmuForceValue"
                                                                                                
    '取得Force格納用変数
    ReDim tmpRetForceVal(CountExistSite)
    
    '最終取得結果用コレクションを準備
    Set tmpForceVal = New Collection
                                                    
    'カンマ区切りのピンリストを配列形式に変換
    Call SeparatePinList(m_PinList, pinNameArr)
    
    'ピンLOOP
    For pinLpCnt = 0 To UBound(pinNameArr)
        'Pinリストからマルチサイト分のチャンネルリストを作成
        Call GetChanList(pinNameArr(pinLpCnt), ALL_SITE, GetChanType(pinNameArr(pinLpCnt)), chListArr)
        
        '現在のモードと設定されているレンジ値をTERADYNE-APIより取得
        TheHdw.APMU.Chans(chListArr).ReadRangesAndMode tmpApmuMode, tmpApmuVrange, tmpApmuIrange
                
        'サイトループ(CHループ)
        For siteLpCnt = 0 To CountExistSite
            '現在の設定モードに応じて電圧レンジ値もしくは電流レンジ値を格納する
            With TheHdw.APMU.Chans(chListArr(siteLpCnt))
                Select Case tmpApmuMode(siteLpCnt)
                    Case apmuForceIMeasureV: 'FIMV
                        status = .ReadForceCurrents(tmpApmuIrange, getForceVal)
                        tmpRetForceVal(siteLpCnt) = getForceVal(0)
                    
                    Case apmuForceVMeasureI: 'FVMI
                        status = .ReadForceVoltages(tmpApmuVrange, getForceVal)
                        tmpRetForceVal(siteLpCnt) = getForceVal(0)
                    
                    Case apmuMeasureV:       'MVモードのときはForce値を0として取り扱う
                        tmpRetForceVal(siteLpCnt) = 0#
                End Select
            End With
        Next siteLpCnt
        'ピン名をキーにForce値の取得結果をコレクションに登録
        Call tmpForceVal.Add(tmpRetForceVal, pinNameArr(pinLpCnt))
    Next pinLpCnt
        
'---------------------------------------------------------------------------
'取得結果算出処理

    'ピンリストの中身が、GangのGPが含まれているとき
    If IsGangPinlist(m_PinList, GetChanType(m_PinList)) = True Then

        'ギャング用のGPがPinListに設定されているときは単一GPかどうかを確認
        Dim gangPinListArr() As String
        Call ConvertStrPinListToArrayPinList(m_PinList, gangPinListArr) 'ピングループを展開せずにカンマ区切り形式の配列に変換

        'Gangのマルチピンには対応できていないのでエラーにする
        If UBound(gangPinListArr) <> 0 Then
            Call MsgBox("Ganged pin group is included in the pinlist.(" & m_PinList & ")  at " & FUNCTION_NAME & "()", vbCritical, FUNCTION_NAME)
            Set GetApmuForceValue = Nothing
            Exit Function
        End If
        
        Dim tmpGangForce As Collection
        Set tmpGangForce = New Collection
        Call tmpGangForce.Add(tmpForceVal.Item(pinNameArr(0)), m_PinList)

        'Gang用取得値を取得結果として返す
        Set GetApmuForceValue = tmpGangForce

    'ピンリストの中身が通常ピンのとき(取得結果をそのまま返す)
    Else
        '取得結果を関数の戻り値としてコレクションで返す
        Set GetApmuForceValue = tmpForceVal
    End If
            
End Function

'Private Function ApmuModeToStr(ByVal ForceMode As ApmuMode) As String
''APMUのMode番号を文字列に変換(未使用)
'
'    Select Case ForceMode
'        Case apmuForceIMeasureV
'            ApmuModeToStr = "FIMV"
'
'        Case apmuForceVMeasureI
'            ApmuModeToStr = "FVMI"
'
'        Case apmuMeasureV
'             ApmuModeToStr = "MV"
'    End Select
'
'End Function

'#Pass-Ok
Private Function ApmuVRangeToValue(ByVal Vrange As ApmuVRange) As Double
'APMUのVRange番号をDouble値に変換

    Select Case Vrange
        Case apmu2V
            ApmuVRangeToValue = 2#
        
        Case apmu5V
            ApmuVRangeToValue = 5#
        
        Case apmu10V
            ApmuVRangeToValue = 10#
        
        Case apmu40V
            ApmuVRangeToValue = 40#
        
        Case apmu35V
            ApmuVRangeToValue = 35#
        
        Case Else
            ApmuVRangeToValue = DC_RANGE_UNKNOWN  '"Unknown"
    End Select

End Function

'#Pass-Ok
Private Function ApmuIRangeToValue(ByVal Irange As ApmuIRange) As Double
'APMUのIRange番号をDouble値に変換

    Select Case Irange
        Case apmu200uA
            ApmuIRangeToValue = 0.0002
        
        Case apmu1mA
            ApmuIRangeToValue = 0.001
        
        Case apmu5mA
            ApmuIRangeToValue = 0.005
        
        Case apmu50mA
            ApmuIRangeToValue = 0.05
        
        Case apmu200nA
            ApmuIRangeToValue = 0.0000002
        
        Case apmu2uA
            ApmuIRangeToValue = 0.000002
        
        Case apmu10uA
            ApmuIRangeToValue = 0.00001
        
        Case apmu40uA
            ApmuIRangeToValue = 0.00004
        
        Case Else
            ApmuIRangeToValue = DC_RANGE_UNKNOWN  '"Unknown"
    End Select

End Function

'######################## 以下 V/I Source I/Fの実装 #############################

'#Pass-Ok
Private Property Get IVISVISource_ForceVRangeList() As String
    IVISVISource_ForceVRangeList = FV_RANGE_LIST
End Property

'#Pass-Ok
Private Property Get IVISVISource_ForceIRangeList() As String
    
    'ピンリストにGangのGPが含まれているとき
    If IsGangPinlist(m_PinList, GetChanType(m_PinList)) = True Then
        IVISVISource_ForceIRangeList = GANG_FI_RANGE_LIST
    Else
        IVISVISource_ForceIRangeList = FI_RANGE_LIST
    End If

End Property

'#Pass-Ok
Private Property Get IVISVISource_MeasVRangeList() As String
    IVISVISource_MeasVRangeList = MV_RANGE_LIST
End Property

'#Pass-Ok
Private Property Get IVISVISource_MeasIRangeList() As String
    
    'ピンリストにGangのGPが含まれているとき
    If IsGangPinlist(m_PinList, GetChanType(m_PinList)) = True Then
        IVISVISource_MeasIRangeList = GANG_MI_RANGE_LIST
    Else
        IVISVISource_MeasIRangeList = MI_RANGE_LIST
    End If
    
End Property

'#Pass-Ok
Private Function IVISVISource_GetForceValue() As Collection

    Set IVISVISource_GetForceValue = Nothing
        
    'FAILサイトが存在するか確認し、存在する場合はエラーとする
    
    If CheckFailSiteExists("GetForceValue") = True Then
        Exit Function
    End If

    Set IVISVISource_GetForceValue = GetApmuForceValue

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

    'ピンリストにGangのGPが含まれているとき
    If IsGangPinlist(m_PinList, GetChanType(m_PinList)) = True Then
                
        'ギャング用のGPがPinListに設定されているときは単一GPかどうかを確認
        Dim pinListArr() As String
        Call ConvertStrPinListToArrayPinList(m_PinList, pinListArr) 'ピングループを展開せずにカンマ区切り形式の配列に変換
        If UBound(pinListArr) <> 0 Then
            Call OutputErrMsg("Ganged pin group is included in the " & m_PinList)
            Exit Function
        End If
                          
        '並列接続数のカウント用にピンリストを展開
        Dim gangMemberPin() As String
        Call SeparatePinList(m_PinList, gangMemberPin)
                        
        '最大設定電流値×並列接続数を、取得結果とする
        tmpLimitVal(0) = FI_LIMIT_LOW * (UBound(gangMemberPin) + 1)
        tmpLimitVal(1) = FI_LIMIT_HIGH * (UBound(gangMemberPin) + 1)
    
        IVISVISource_GetForceILimit = tmpLimitVal
        
    'ピンリストの中身が通常ピンのとき
    Else
        tmpLimitVal(0) = FI_LIMIT_LOW
        tmpLimitVal(1) = FI_LIMIT_HIGH
        IVISVISource_GetForceILimit = tmpLimitVal
    End If

End Function

'#Pass-Ok
Private Function IVISVISource_GetMeasureRange() As Collection
    
    Set IVISVISource_GetMeasureRange = Nothing
    
    'FAILサイトが存在するか確認し、存在する場合はエラーとする
    If CheckFailSiteExists("GetMeasureRange") = True Then
        Exit Function
    End If

    Set IVISVISource_GetMeasureRange = GetApmuMeasureRange

End Function

'#Pass-Ok
Private Property Get IVISVISource_BoardName() As String

    IVISVISource_BoardName = BOARD_NAME

End Property

'#Pass-Ok
Private Property Let IVISVISource_PinList(ByVal PinList As String)
    
    m_PinList = PinList

End Property

'#Pass-Ok
Private Property Get IVISVISource_PinList() As String
    
    IVISVISource_PinList = m_PinList

End Property

'#Pass-Ok
Private Sub IVISVISource_SetFVMI(ByVal ForceV As Variant, ByVal ClampI As Double, Optional ByVal site As Long = -1&, Optional ByVal ConnectOn As Boolean = True)

    Call SetFVMI(m_PinList, ForceV, Abs(ClampI), site, , , ConnectOn)

End Sub
'#V21-Release
Private Sub IVISVISource_SetFVMIMulti(ByVal ForceV As Variant, ByVal ClampI As Double, Optional ByVal ConnectOn As Boolean = True)

    Call SetFVMIMulti(m_PinList, ForceV, Abs(ClampI), , , ConnectOn)

End Sub
'#Pass-Ok
Private Sub IVISVISource_SetFIMV(ByVal ForceI As Variant, ByVal ClampV As Double, Optional ByVal site As Long = -1&, Optional ByVal ConnectOn As Boolean = True)

    Call SetFIMV(m_PinList, ForceI, Abs(ClampV), site, , , ConnectOn)

End Sub
'#V21-Release
Private Sub IVISVISource_SetFIMVMulti(ByVal ForceI As Variant, ByVal ClampV As Double, Optional ByVal ConnectOn As Boolean = True)

    Call SetFIMVMulti(m_PinList, ForceI, Abs(ClampV), , , ConnectOn)

End Sub
'#Pass-Ok
Private Sub IVISVISource_SetMV(ByVal ClampV As Double, Optional ByVal site As Long = -1&, Optional ByVal ConnectOn As Boolean = True)
    
    Call SetMV(m_PinList, ClampV, site, , ConnectOn)

End Sub
'#V21-Release
Private Sub IVISVISource_SetMVMulti(ByVal ClampV As Double, Optional ByVal ConnectOn As Boolean = True)
    
    Call SetMVMulti(m_PinList, ClampV, , ConnectOn)

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
    
    Call ChangeMIRange(m_PinList, Abs(ClampI), site)

End Sub
'#V21-Release
Private Sub IVISVISource_ChangeMIRangeMulti(ByVal ClampI As Double)

    Call ChangeMIRangeMulti(m_PinList, Abs(ClampI))

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
    
    status = ReadMultiResult(m_PinList, retResult, m_ResultsV_APMU)

    If status = False Then
        Call OutputErrMsg("Not Found Result of " & m_PinList & " at ReadMVMultiResult()")
        Stop
    End If

End Sub

'#Pass-Ok
Private Sub IVISVISource_ReadMIMultiResult(retResult() As Double)
    
    Dim status As Boolean
    
    status = ReadMultiResult(m_PinList, retResult, m_ResultsI_APMU)

    If status = False Then
        Call OutputErrMsg("Not Found Result of " & m_PinList & " at ReadMIMultiResult()")
        Stop
    End If

End Sub

'################################################################################


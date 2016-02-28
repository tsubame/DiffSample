VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVISHdvis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'概要:
'   HDVISボード制御用クラス
'
'目的:
'   Ⅰ:HDVISボードの使い勝手向上
'   Ⅱ:V/ISourceとしてのI/F提供
'
'作成者:
'   SLSI今手
'
'XLibSTD_CommonDCModよりHDVIS関連機能を整理
'
'注意事項:
'利用時には、条件付コンパイル引数にHDVIS_USE=1の記述が必要です。
'
'HDVIS電流測定不具合対応版:
'HDVISボードハードウェア側での対応が完了するまでの暫定対策です。
'
'
'Code Checked
'Comment Checked
'

Option Explicit

'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const HDVIS_USE = 0      'HVDISリソースの使用　 0：未使用、0以外：使用  <PALSとEeeAuto両方で使用>

#If HDVIS_USE <> 0 Then
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
Private Const BOARD_NAME = "dcHDVIS"

Private Const FV_LIMIT_HIGH = 10# * V
Private Const FV_LIMIT_LOW = -10# * V
Private Const FI_LIMIT_HIGH = 200# * mA
Private Const FI_LIMIT_LOW = -200# * mA

Private Const FV_RANGE_LIST = "10V"
Private Const MV_RANGE_LIST = "10V"
Private Const FI_RANGE_LIST = "5uA,50uA,500uA,5mA,50mA,200mA"
Private Const MI_RANGE_LIST = "5uA,50uA,500uA,5mA,50mA,200mA"

Private m_PinList As String
Private m_ResultsV_HDVIS As Collection
Private m_ResultsI_HDVIS As Collection


'--- 不具合対応用に追加 --------------------
Private Const DEF_PERIOD = 0.000005 '5uA
Private m_ParamObj As New CVISHdvisParameter
'-------------------------------------------

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
    Optional ByVal FVRange As HdvisVRange = DEF_RANGE, Optional ByVal MIRange As HdvisIRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'    HDVISを電圧印加状態に設定する｡
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
'    ■PinListに対応するHDVISを電圧印加状態にする｡
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
    Dim FVRangeBySite() As HdvisVRange

    'Error Check ************************************
    Const FunctionName = "SetFVMI_HDVIS"
    'Pass
    If CheckPinList(PinList, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceV, -10 * V, 10 * V, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 0 * mA, 200 * mA * 4, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    Call ConvertVariableToArray(ForceV)

    ReDim FVRangeBySite(CountExistSite)
    For curSite = 0 To CountExistSite
        'Pass
        If FVRange = DEF_RANGE Then
            'Pass
            FVRangeBySite(curSite) = GetHdvisFVRange(ForceV(curSite))
        Else
            'Pass
            FVRangeBySite(curSite) = FVRange
        End If
    Next curSite
    
    If MIRange = DEF_RANGE Then
        'Pass
        MIRange = GetHdvisMIRange(ClampI)
    End If
    
    TheHdw.HDVIS.RelayMode = hdvisRelayModeHotSwitch
    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        'Pass
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
        
            'Main Part ---------------------------------------
            'Pass
            With TheHdw.HDVIS.Pins(PinList)
                .alarm(hdvisAlarmOverLoad) = False
                Call .ModeFVMI(MIRange)
                .ClampCurrent(MIRange) = ClampI
                .ForceVoltage(FVRangeBySite(curSite)) = ForceV(curSite)
                .Gate = True
            End With
            '-------------------------------------------------
            
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

    If ConnectOn = True Then
        'Pass
        Call ConnectPins(PinList, site)
    End If
    
End Sub
'#V21-Release
Private Sub SetFVMIMulti( _
    ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, _
    Optional ByVal FVRange As HdvisVRange = DEF_RANGE, Optional ByVal MIRange As HdvisIRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'    HDVISを電圧印加状態に設定する｡ (Site同時)
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
'    ■PinListに対応するHDVISを電圧印加状態にする｡
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
    Dim FVRangeBySite() As HdvisVRange

    'Error Check ************************************
    Const FunctionName = "SetFVMIMulti_HDVIS"
    If CheckPinList(PinList, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceV, -10 * V, 10 * V, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 0 * mA, 200 * mA * 4, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
     'site0only設定************************************
    Call ConvertVariableToArray(ForceV)

    ReDim FVRangeBySite(1)
     curSite = 0
     
        If FVRange = DEF_RANGE Then
           
            FVRangeBySite(curSite) = GetHdvisFVRange(ForceV(curSite))
        Else
           
            FVRangeBySite(curSite) = FVRange
        End If
    '************************************************
    
    If MIRange = DEF_RANGE Then
 
        MIRange = GetHdvisMIRange(ClampI)
    End If
    
    TheHdw.HDVIS.RelayMode = hdvisRelayModeHotSwitch

        
            'Main Part ---------------------------------------
          
            With TheHdw.HDVIS.Pins(PinList)
                .alarm(hdvisAlarmOverLoad) = False
                Call .ModeFVMI(MIRange)
                .ClampCurrent(MIRange) = ClampI
                .ForceVoltage(FVRangeBySite(curSite)) = ForceV(curSite)
                .Gate = True
            End With
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
    Optional ByVal FIRange As HdvisIRange = DEF_RANGE, Optional ByVal MVRange As HdvisVRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'    HDVISを電流印加状態に設定する｡
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
'    ■PinListに対応するHDVISを電流印加状態にする｡
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
    Dim FIRangeBySite() As HdvisIRange

    'Error Check ************************************
    Const FunctionName = "SetFIMV_HDVIS"
    'Pass
    If CheckPinList(PinList, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceI, -200 * mA * 4, 200 * mA * 4, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampV, 0 * V, 10 * V, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************

    Call ConvertVariableToArray(ForceI)

    ReDim FIRangeBySite(CountExistSite)
    For curSite = 0 To CountExistSite
        'Pass
        If FIRange = DEF_RANGE Then
            'Pass
            FIRangeBySite(curSite) = GetHdvisFIRange(ForceI(curSite))
        Else
            'Pass
            FIRangeBySite(curSite) = FIRange
        End If
    Next curSite

    If MVRange = DEF_RANGE Then
        'Pass
        MVRange = GetHdvisMVRange(ClampV)
    End If

    TheHdw.HDVIS.RelayMode = hdvisRelayModeHotSwitch
    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        'Pass
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then

            'Main Part ---------------------------------------
            'Pass
            With TheHdw.HDVIS.Pins(PinList)
                .alarm(hdvisAlarmOverLoad) = False
                Call .ModeFIMV(MVRange)
                .ClampVoltage(MVRange) = ClampV
                .ForceCurrent(FIRangeBySite(curSite)) = ForceI(curSite)
                .Gate = True
            End With
            '-------------------------------------------------

        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

    If ConnectOn = True Then
        'Pass
        Call ConnectPins(PinList, site)
    End If

End Sub
'#V21-Release
Private Sub SetFIMVMulti( _
    ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, _
    Optional ByVal FIRange As HdvisIRange = DEF_RANGE, Optional ByVal MVRange As HdvisVRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'    HDVISを電流印加状態に設定する｡(Site同時)
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
'    ■PinListに対応するHDVISを電流印加状態にする｡
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
    Dim FIRangeBySite() As HdvisIRange

    'Error Check ************************************
    Const FunctionName = "SetFIMVMulti_HDVIS"

    If CheckPinList(PinList, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceI, -200 * mA * 4, 200 * mA * 4, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampV, 0 * V, 10 * V, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************

'site0only設定************************************
    Call ConvertVariableToArray(ForceI)

    ReDim FIRangeBySite(1)
    curSite = 0
        If FIRange = DEF_RANGE Then
            FIRangeBySite(curSite) = GetHdvisFIRange(ForceI(curSite))
        Else
            FIRangeBySite(curSite) = FIRange
        End If
'************************************************

    If MVRange = DEF_RANGE Then
        MVRange = GetHdvisMVRange(ClampV)
    End If

    TheHdw.HDVIS.RelayMode = hdvisRelayModeHotSwitch

            'Main Part ---------------------------------------
            
            With TheHdw.HDVIS.Pins(PinList)
                .alarm(hdvisAlarmOverLoad) = False
                Call .ModeFIMV(MVRange)
                .ClampVoltage(MVRange) = ClampV
                .ForceCurrent(FIRangeBySite(curSite)) = ForceI(curSite)
                .Gate = True
            End With
            '-------------------------------------------------

    
       If ConnectOn = True Then
        Call ConnectPinsMultiPins(PinList)
'       Call ConnectPinsMultiChans(pinList)
       End If
       
End Sub
'#Pass-Ok
Private Sub SetMV( _
    ByVal PinList As String, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal MVRange As HdvisVRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'    HDVISを電圧測定状態(無負荷)に設定する。
'
'パラメータ:
'    [PinList]    In  対象ピンリスト。
'    [Site]       In  サイト指定。オプション(Default -1)
'    [MVRange]    In  電圧測定レンジ。オプション(Default -1)
'    [ConnectOn]  In  デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■電圧クランプの機能はありません
'    ■PinListに対応するHDVISを無負荷電圧測定状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    
    'Error Check ************************************
    Const FunctionName = "SetMV_HDVIS"
    'Pass
    If CheckPinList(PinList, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************
    
    If MVRange = DEF_RANGE Then
        MVRange = HDVIS10V
    End If
    
    TheHdw.HDVIS.RelayMode = hdvisRelayModeHotSwitch
    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        'Pass
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then

            'Main Part ---------------------------------------
            'Pass
            With TheHdw.HDVIS.Pins(PinList)
                .alarm(hdvisAlarmOverLoad) = False
                Call .ModeMV(MVRange)
                .Gate = True
            End With
            '-------------------------------------------------

        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

    If ConnectOn = True Then
        'Pass
        Call ConnectPins(PinList, site)
    End If

End Sub
'#V21-Release
Private Sub SetMVMulti( _
    ByVal PinList As String, _
   Optional ByVal MVRange As HdvisVRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)

'内容:
'    HDVISを電圧測定状態(無負荷)に設定する。(Site同時)
'
'パラメータ:
'    [PinList]    In  対象ピンリスト。
'    [MVRange]    In  電圧測定レンジ。オプション(Default -1)
'    [ConnectOn]  In  デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■電圧クランプの機能はありません
'    ■PinListに対応するHDVISを無負荷電圧測定状態にする｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long
    
    'Error Check ************************************
    Const FunctionName = "SetMVMulti_HDVIS"

    If CheckPinList(PinList, chHDVIS, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
    If MVRange = DEF_RANGE Then
        MVRange = HDVIS10V
    End If
    
    TheHdw.HDVIS.RelayMode = hdvisRelayModeHotSwitch

            'Main Part ---------------------------------------

            With TheHdw.HDVIS.Pins(PinList)
                .alarm(hdvisAlarmOverLoad) = False
                Call .ModeMV(MVRange)
                .Gate = True
            End With
            '-------------------------------------------------

       If ConnectOn = True Then
        Call ConnectPinsMultiPins(PinList)
'       Call ConnectPinsMultiChans(pinList)
       End If
       
End Sub
'#Pass-Ok
Private Sub DisconnectPins(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)
'内容:
'    HDVISをデバイスから切り離す｡
'
'パラメータ:
'    [PinList]  In  対象ピンリスト。
'    [Site]     In  サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するHDVISをデバイスから切り離す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■切り離した後、0V印加状態に設定する(接続はしない)。
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long

    'Error Check ************************************
    Const FunctionName = "DisconnectPins_HDVIS"
    'Pass
    If CheckPinList(PinList, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************

    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        'Pass
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then

            'Main Part ---------------------------------------
            'Pass
            With TheHdw.HDVIS.Pins(PinList)
                .relay = False
                .Gate = False
                Call .ModeFVMI(hdvis200mA)
                .ForceVoltage(HDVIS10V) = 0 * V
                .ClampCurrent(hdvis200mA) = 200 * mA
                .alarm(hdvisAlarmOverLoad) = True
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
'    HDVISをデバイスから切り離す｡(Site同時)
'
'パラメータ:
'    [PinList]  In  対象ピンリスト。

'
'戻り値:
'
'注意事項:
'    ■PinListに対応するHDVISをデバイスから切り離す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■切り離した後、0V印加状態に設定する(接続はしない)。
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long

    'Error Check ************************************
    Const FunctionName = "DisconnectPinsMulti_HDVIS"
    If CheckPinList(PinList, chHDVIS, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************
    
   'Main Part ---------------------------------------

   With TheHdw.HDVIS.Pins(PinList)
   .relay = False
   .Gate = False
   Call .ModeFVMI(hdvis200mA)
   .ForceVoltage(HDVIS10V) = 0 * V
   .ClampCurrent(hdvis200mA) = 200 * mA
   .alarm(hdvisAlarmOverLoad) = True
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
            Call GetChanList(TmpPinList, curSite, chHDVIS, Channels)
            For chan = 0 To UBound(Channels)
                TheHdw.HDVIS.Chans(Channels(chan)).MergedChannels = False
            Next
        End If
    Next curSite
    
End Sub

'#Pass-Ok
Private Sub ConnectPins(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)
'内容:
'   HDVISをデバイスに接続する｡
'
'パラメータ:
'    [PinList]  In  対象ピンリスト。
'    [Site]     In  サイト指定。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するHDVISをデバイスに接続する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'
    Dim siteStatus As LoopStatus
    Dim curSite As Long

    'Error Check ************************************
    Const FunctionName = "ConnectPins_HDVIS"
    'Pass
    If CheckPinList(PinList, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************

    'Site Loop **************************************
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        'Pass
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then

            'Main Part ---------------------------------------
            'Pass
            TheHdw.HDVIS.Pins(PinList).relay = True
            '-------------------------------------------------

        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

End Sub
'#V21-Release
Private Sub ConnectPinsMultiPins(ByVal PinList As String)
'内容:
'   HDVISをデバイスに接続する｡ (Site同時)
'
'パラメータ:
'    [PinList]  In  対象ピンリスト。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するHDVISをデバイスに接続する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long

    'Error Check ************************************
    Const FunctionName = "ConnectPinsMulti_HDVIS"

    If CheckPinList(PinList, chHDVIS, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************


            'Main Part ---------------------------------------
            TheHdw.HDVIS.Pins(PinList).relay = True
            '-------------------------------------------------


End Sub
'#V21-Release
Private Sub ConnectPinsMultiChans(ByVal PinList As String)
'内容:
'   HDVISをデバイスに接続する｡(Site同時)
'
'パラメータ:
'    [PinList]  In  対象ピンリスト。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するHDVISをデバイスに接続する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡
'
'    Dim SiteStatus As loopStatus
'    Dim CurSite As Long
     Dim Channels() As Long

    'Error Check ************************************
    Const FunctionName = "ConnectPinsMulti_HDVIS"

    If CheckPinList(PinList, chHDVIS, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************


            'Main Part ---------------------------------------
            TheHdw.HDVIS.Chans(Channels).relay = True
            '-------------------------------------------------


End Sub
'#Pass-Ok
Private Sub SetGND( _
    ByVal PinList As String, Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True _
)
'内容:
'    HDVISを接地する｡
'
'パラメータ:
'    [PinList]    In  対象ピンリスト。
'    [Site]       In  サイト指定。オプション(Default -1)
'    [ConnectOn]  In  デバイスに接続するかどうか。オプション(Default True)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するHDVISを接地する｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■Siteで対象のサイトを指定。-1(デフォルト)を入れると、全サイトが対象。
'    ■非アクティブサイトに対しては何もしない｡
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)
'
    'Pass
    Call SetFVMI(PinList, 0 * V, 200 * mA, site, , , ConnectOn)

End Sub

'#Pass-Ok
Private Sub ChangeMIRange( _
    ByVal PinList As String, ByVal ClampI As Double, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal MIRange As HdvisIRange = DEF_RANGE _
)
'内容:
'   電圧印加状態のHDVISの電流測定レンジを変更する｡
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
'    ■PinListに対応するHDVISの印加電圧を変えず､電流測定レンジを変更する｡
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

    Dim CurMode() As HdvisMode
    Dim CurVRange() As HdvisVRange
    Dim CurIRange() As HdvisIRange

    Dim ErrChans As String

    'Error Check ************************************
    Const FunctionName = "ChangeMIRange_HDVIS"
    'Pass
    If CheckPinList(PinList, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 0 * mA, 200 * mA * 4, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************

    If MIRange = DEF_RANGE Then
        'Pass
        MIRange = GetHdvisMIRange(ClampI)
    End If

    TheHdw.HDVIS.RelayMode = hdvisRelayModeHotSwitch
    'Site Loop **************************************
    ErrChans = ""
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        'Pass
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then

            'Main Part ---------------------------------------
            'Pass
            Call GetChanList(PinList, curSite, chHDVIS, Channels)
            For chan = 0 To UBound(Channels)
                'Pass
                With TheHdw.HDVIS.Chans(Channels(chan))
                    Call .ReadRangesAndMode(CurMode, CurVRange, CurIRange)
                    If CurMode(0) <> hdvisForceVMeasureI Then
                        'Pass
                        ErrChans = ErrChans & Channels(chan) & ","
'                        Call OutputErrMsg("HDVIS(" & Channels(chan) & ") is not MI Mode. (at ChangeMIRange_HDVIS)")
                    Else
                        'Pass
                        .ClampCurrent(MIRange) = ClampI
                    End If
                End With
            Next chan
            '-------------------------------------------------

        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

    If ErrChans <> "" Then
        Call OutputErrMsg("HDVIS(" & ErrChans & ") is not MI Mode. (at ChangeMIRange_HDVIS)")
    End If

End Sub
'#V21-Release
Private Sub ChangeMIRangeMulti( _
    ByVal PinList As String, ByVal ClampI As Double, _
    Optional ByVal MIRange As HdvisIRange = DEF_RANGE _
)

'内容:
'   電圧印加状態のHDVISの電流測定レンジを変更する｡(Site同時)
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'    [ClampI]    In   クランプ電流値。
'    [MIRange]   In   電流測定レンジ。オプション(Default -1)
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するHDVISの印加電圧を変えず､電流測定レンジを変更する｡
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

    Dim CurMode() As HdvisMode
    Dim CurVRange() As HdvisVRange
    Dim CurIRange() As HdvisIRange

    Dim ErrChans As String

    'Error Check ************************************
    Const FunctionName = "ChangeMIRangeMulti_HDVIS"
    
    If CheckPinList(PinList, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 0 * mA, 200 * mA * 4, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************

    If MIRange = DEF_RANGE Then
        MIRange = GetHdvisMIRange(ClampI)
    End If

    TheHdw.HDVIS.RelayMode = hdvisRelayModeHotSwitch

    ErrChans = ""

            'Main Part ---------------------------------------
            Call GetActiveChanList(PinList, chHDVIS, Channels)
           
                With TheHdw.HDVIS.Chans(Channels)
                    Call .ReadRangesAndMode(CurMode, CurVRange, CurIRange)
                    If CurMode(0) <> hdvisForceVMeasureI Then
                        ErrChans = "AnyChans"
'                        Call OutputErrMsg("HDVIS(" & Channels(chan) & ") is not MI Mode. (at ChangeMIRange_HDVIS)")
                    Else
                        .ClampCurrent(MIRange) = ClampI
                    End If
                End With
   
            '-------------------------------------------------


    If ErrChans <> "" Then
        Call OutputErrMsg("HDVIS(" & ErrChans & ") is not MI Mode. (at ChangeMIRangeMulti_HDVIS)")
    End If
    
End Sub
'#Pass-Ok
Private Sub MeasureV( _
    ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal UseLPF As Boolean = False _
)
'内容:
'    HDVISの電圧値を測定する｡
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
'    ■PinNameに対応するHDVISの電圧値を測定する｡
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
    Dim chan As Long
    Dim Samples() As Double
    Dim n As Long

    Dim CurMode() As HdvisMode
    Dim CurVRange() As HdvisVRange
    Dim CurIRange() As HdvisIRange
    
    Dim ErrPins As String

    'Error Check ************************************
    Const FunctionName = "MeasureV_HDVIS"
    'Pass
    If CheckPinList(PinName, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If CheckSinglePins(PinName, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If CheckResultArray(retResult, FunctionName) = False Then Stop: Exit Sub
    If CheckAvgNum(avgNum, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************

    If site = ALL_SITE Then
        'Pass
        For curSite = 0 To CountExistSite
            'Pass
            retResult(curSite) = 0
        Next curSite
    Else
        'Pass
        If IsActiveSite(site) = False Then
            'Pass
            retResult(site) = 0
            Exit Sub
        End If
    End If

    'Site Loop **************************************
    ErrPins = ""
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        'Pass
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then

            'Main Part ---------------------------------------
            'Pass
            Call GetChanList(PinName, curSite, chHDVIS, Channels)

            With TheHdw.HDVIS.Pins(PinName)
                Call .ReadRangesAndMode(CurMode, CurVRange, CurIRange)
                If CurMode(0) = hdvisForceVMeasureI Then
                    'Pass
                    ErrPins = ErrPins & PinName & "(site" & curSite & "),"
'                    Call OutputErrMsg(PinName & " is not MV Mode (at MeasureV_HDVIS)")
                Else
                    'Pass
                    .LowPassFilter = UseLPF
                    Call .measure(avgNum, Samples)
                    retResult(curSite) = Samples(0)
                End If
            End With
            '-------------------------------------------------

        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************
    
    If ErrPins <> "" Then
        Call OutputErrMsg(ErrPins & " is not MV Mode (at MeasureV_HDVIS)")
    End If

End Sub

'#Pass-Ok
Private Sub MeasureI( _
    ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal UseLPF As Boolean = False _
)
'内容:
'   HDVISの電流値を測定する｡
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
'    ■PinNameに対応するHDVISの電流値を測定する｡
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
    Dim chan As Long
    Dim Samples() As Double
    Dim n As Long

    Dim CurMode() As HdvisMode
    Dim CurVRange() As HdvisVRange
    Dim CurIRange() As HdvisIRange

    Dim ErrPins As String

    'Error Check ************************************
    Const FunctionName = "MeasureI_HDVIS"
    'Pass
    If CheckPinList(PinName, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If CheckSinglePins(PinName, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If CheckResultArray(retResult, FunctionName) = False Then Stop: Exit Sub
    If CheckAvgNum(avgNum, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************

    If site = ALL_SITE Then
        'Pass
        For curSite = 0 To CountExistSite
            'Pass
            retResult(curSite) = 0
        Next curSite
    Else
        'Pass
        If IsActiveSite(site) = False Then
            'Pass
            retResult(site) = 0
            Exit Sub
        End If
    End If

    'Site Loop **************************************
    ErrPins = ""
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        'Pass
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
            'Main Part ---------------------------------------
            'Pass
            Call GetChanList(PinName, curSite, chHDVIS, Channels)
            'Pass
            With TheHdw.HDVIS.Pins(PinName)
                Call .ReadRangesAndMode(CurMode, CurVRange, CurIRange)
                If CurMode(0) <> hdvisForceVMeasureI Then
                    'Pass
                    ErrPins = ErrPins & PinName & "(site" & curSite & "),"
'                    Call OutputErrMsg(PinName & " is not MI Mode (at MeasureI_HDVIS)")
                Else
                    'Pass
                    .LowPassFilter = UseLPF
                    'マージ接続ピンのメジャー時にはSlewRateをSlowに設定する（マニュアルP165より）
                    If IsGangPinlist(PinName, GetChanType(PinName)) = True Then
                        .SlewRate = hdvisSlewRateSlow
                        Call .measure(avgNum, Samples)
                        .SlewRate = hdvisSlewRateNormal
                    Else
                        '最適化パラメータによるメジャー
                        If (CurIRange(0) = hdvis5uA) Or (CurIRange(0) = hdvis50uA) Then
                            'HDVISパラメータ最適化の設定
                            TheHdw.HDVIS.Period = m_ParamObj.Period(PinName, CurIRange(0)) 'Period値を設定
                            .SlewRate = m_ParamObj.SlewRate                                'SlewRate値を設定
                            
                            'デバック用設定条件確認用
                            'Call m_ParamObj.DispMeasureCondition(PinName)
                            Call .measure(m_ParamObj.Samples(PinName, CurIRange(0)), Samples) 'Samples値設定とMeasure
                                                        
                            'PeriodとSlewRateの設定値をデフォルト値に設定する
                            TheHdw.HDVIS.Period = DEF_PERIOD
                            .SlewRate = hdvisSlewRateNormal
                        Else
                            '最適化パラメータを使用しないメジャー(従来処理)
                            Call .measure(avgNum, Samples)
                        End If
                    End If
                    retResult(curSite) = Samples(0)
                End If
            End With
            '-------------------------------------------------
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************
    
    If ErrPins <> "" Then
        Call OutputErrMsg(ErrPins & " is not MI Mode (at MeasureI_HDVIS)")
    End If

End Sub

'#Pass-Ok
Private Sub MeasureI_Org( _
    ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal UseLPF As Boolean = False _
)
'内容:
'   HDVISの電流値を測定する｡
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
'    ■PinNameに対応するHDVISの電流値を測定する｡
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
    Dim chan As Long
    Dim Samples() As Double
    Dim n As Long

    Dim CurMode() As HdvisMode
    Dim CurVRange() As HdvisVRange
    Dim CurIRange() As HdvisIRange

    Dim ErrPins As String

    'Error Check ************************************
    Const FunctionName = "MeasureI_HDVIS"
    'Pass
    If CheckPinList(PinName, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If CheckSinglePins(PinName, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If CheckResultArray(retResult, FunctionName) = False Then Stop: Exit Sub
    If CheckAvgNum(avgNum, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************

    If site = ALL_SITE Then
        'Pass
        For curSite = 0 To CountExistSite
            'Pass
            retResult(curSite) = 0
        Next curSite
    Else
        'Pass
        If IsActiveSite(site) = False Then
            'Pass
            retResult(site) = 0
            Exit Sub
        End If
    End If

    'Site Loop **************************************
    ErrPins = ""
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        'Pass
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then
            'Main Part ---------------------------------------
            'Pass
            Call GetChanList(PinName, curSite, chHDVIS, Channels)
            'Pass
            With TheHdw.HDVIS.Pins(PinName)
                Call .ReadRangesAndMode(CurMode, CurVRange, CurIRange)
                If CurMode(0) <> hdvisForceVMeasureI Then
                    'Pass
                    ErrPins = ErrPins & PinName & "(site" & curSite & "),"
'                    Call OutputErrMsg(PinName & " is not MI Mode (at MeasureI_HDVIS)")
                Else
                    'Pass
                    .LowPassFilter = UseLPF
                    'マージ接続ピンのメジャー時にはSlewRateをSlowに設定する（マニュアルP165より）
                    If IsGangPinlist(PinName, GetChanType(PinName)) = True Then
                        .SlewRate = hdvisSlewRateSlow
                        Call .measure(avgNum, Samples)
                        .SlewRate = hdvisSlewRateNormal
                    Else
                        Call .measure(avgNum, Samples)
                    End If
                    retResult(curSite) = Samples(0)
                End If
            End With
            '-------------------------------------------------
        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************
    
    If ErrPins <> "" Then
        Call OutputErrMsg(ErrPins & " is not MI Mode (at MeasureI_HDVIS)")
    End If

End Sub



'#Pass-Ok
Private Sub MeasureVMulti(ByVal PinList As String, ByVal avgNum As Long, Optional ByVal UseLPF As Boolean = False)
'内容:
'   HDVISの電圧値を測定する｡
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

    Dim CurMode() As HdvisMode
    Dim CurVRange() As HdvisVRange
    Dim CurIRange() As HdvisIRange
    Dim chan As Long


    'Error Check ************************************
    Const FunctionName = "MeasureVMulti_HDVIS"
    'Pass
    If CheckPinList(PinList, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If CheckAvgNum(avgNum, FunctionName) = False Then Stop: Exit Sub
    '************************************************

    'Check Ganged Pins ********************************************
    Call GetActiveChanList(PinList, chHDVIS, Channels)
    Call SeparatePinList(PinList, pinNames)
    If (UBound(pinNames) + 1) * CountActiveSite <> UBound(Channels) + 1 Then
        'Pass
        Call OutputErrMsg(PinList & " Including Ganged Pins. (at MeasureVMulti_HDVIS)")
        Exit Sub
    End If
    '**************************************************************

    'Check MV Mode ************************************************
    Call TheHdw.HDVIS.Pins(PinList).ReadRangesAndMode(CurMode, CurVRange, CurIRange)
    For chan = 0 To UBound(CurMode)
        'Pass
        If CurMode(chan) = hdvisForceVMeasureI Then
            'Pass
            Call OutputErrMsg(PinList & " Including not MV Mode Pins. (at MeasureVMulti_HDVIS)")
            Exit Sub
        End If
    Next chan
    '**************************************************************

    'Measurement **************************************************
    With TheHdw.HDVIS.Pins(PinList)
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
        'Pass
        For curSite = 0 To CountExistSite
            'Pass
            If IsActiveSite(curSite) Then
                'Pass
                PerPinResults(Pin)(curSite) = Samples(i)
                i = i + 1
           End If
        Next curSite
    Next Pin

    '---------------------------------------------------------

    'Average & Regist ----------------------------------------
    Set m_ResultsV_HDVIS = New Collection
    For Pin = 0 To UBound(pinNames)
        'Pass
        Call m_ResultsV_HDVIS.Add(PerPinResults(Pin), pinNames(Pin))
    Next Pin
    '---------------------------------------------------------
    '**************************************************************

End Sub

'#Pass-Ok
Private Sub MeasureIMulti(ByVal PinList As String, ByVal avgNum As Long, Optional ByVal UseLPF As Boolean = False)
'内容:
'   HDVISの電流値を測定する｡
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

    Dim CurMode() As HdvisMode
    Dim CurVRange() As HdvisVRange
    Dim CurIRange() As HdvisIRange
    Dim chan As Long


    'Error Check ************************************
    Const FunctionName = "MeasureIMulti_HDVIS"
    'Pass
    If CheckPinList(PinList, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If CheckAvgNum(avgNum, FunctionName) = False Then Stop: Exit Sub
    '************************************************

    'Check Ganged Pins ********************************************
    Call GetActiveChanList(PinList, chHDVIS, Channels)
    Call SeparatePinList(PinList, pinNames)
    If (UBound(pinNames) + 1) * CountActiveSite <> UBound(Channels) + 1 Then
        'Pass
        Call OutputErrMsg(PinList & " Including Ganged Pins. (at MeasureIMulti_HDVIS)")
        Exit Sub
    End If
    '**************************************************************

    'Check MI Mode ************************************************
    Call TheHdw.HDVIS.Pins(PinList).ReadRangesAndMode(CurMode, CurVRange, CurIRange)
    For chan = 0 To UBound(CurMode)
        'Pass
        If CurMode(chan) <> hdvisForceVMeasureI Then
            'Pass
            Call OutputErrMsg(PinList & " Including not MI Mode Pins. (at MeasureIMulti_HDVIS)")
            Exit Sub
        End If
    Next chan
    '**************************************************************

    'Measurement **************************************************
    With m_ParamObj
        '5uAもしくは50uAに設定されているピンがメジャーPinGPに存在するか確認
        If .ChkIRangeIncluded(hdvis5uA, CurIRange) Or .ChkIRangeIncluded(hdvis50uA, CurIRange) Then
                        
            '$$$ 同一パラメータが適用できるかのCheck処理 $$$$$$$$$$$$$$$$$$$$$
            'すべてのピンのレンジ設定値が同一であることを確認
            If .ChkSameIrange(CurIRange) <> True Then
                Call OutputErrMsg(PinList & " Irange parameter is not the same")
            End If
            'すべてのピンのメジャーパラメータが同一であることを確認(SamplesとPeriod値の確認)
            If .ChkSameMeasureParameter(PinList, CurIRange(0)) <> True Then
                Call OutputErrMsg(PinList & " measure parameter is not the same")
            End If
            '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
                                                        
            'メジャーPinGpの先頭ピンの状態を渡してメジャーパラメータを使用してメジャー
            'チェック処理によりメジャーPinGPのパラメータは同一であるので、
            'GP先頭ピンのパラメータを適用して Measureの処理を実行する。
            
            'ピリオド値取得と設定(Gp先頭ピン名とIrange設定よりPeriodの値を取得し設定)
            TheHdw.HDVIS.Period = .Period(pinNames(0), CurIRange(0))
            
            With TheHdw.HDVIS.Pins(PinList)
                
                'SlewRate取得と設定（ワークシートより値取得）
                .SlewRate = m_ParamObj.SlewRate
                .LowPassFilter = UseLPF
                
                'debug用状態表示 （将来的にはココにスナップショットを埋め込むか）
                'Call m_ParamObj.DispMeasureCondition(PinList)
                
                'Measure実行(Gp先頭ピン名とIrange設定よりSamplesの値を取得し設定)
                Call .measure(m_ParamObj.Samples(pinNames(0), CurIRange(0)), Samples) 'Sample数取得と設定
                
                'PeriodとSlewRateの設定値をデフォルト値に設定する
                TheHdw.HDVIS.Period = DEF_PERIOD
                .SlewRate = hdvisSlewRateNormal
            End With
        
        Else
            'メジャーパラメータ適用外のメジャー
            With TheHdw.HDVIS.Pins(PinList)
                .LowPassFilter = UseLPF
                Call .measure(avgNum, Samples)
            End With
        End If
    End With
    '**************************************************************

    'Regist Results ***********************************************
    'Initialize ----------------------------------------------
    PerPinResults = CreateEmpty2DArray(UBound(pinNames), CountExistSite)
    '---------------------------------------------------------

    'Summation -----------------------------------------------
    i = 0
    For Pin = 0 To UBound(pinNames)
        'Pass
        For curSite = 0 To CountExistSite
            'Pass
            If IsActiveSite(curSite) Then
                'Pass
                PerPinResults(Pin)(curSite) = Samples(i)
                i = i + 1
            End If
        Next curSite
    Next Pin

    '---------------------------------------------------------

    'Average & Regist ----------------------------------------
    Set m_ResultsI_HDVIS = New Collection
    For Pin = 0 To UBound(pinNames)
        'Pass
        Call m_ResultsI_HDVIS.Add(PerPinResults(Pin), pinNames(Pin))
    Next Pin
    '---------------------------------------------------------
    '**************************************************************

End Sub

'#Pass-Ok
Private Sub MeasureIMulti_Org(ByVal PinList As String, ByVal avgNum As Long, Optional ByVal UseLPF As Boolean = False)
'内容:
'   HDVISの電流値を測定する｡
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

    Dim CurMode() As HdvisMode
    Dim CurVRange() As HdvisVRange
    Dim CurIRange() As HdvisIRange
    Dim chan As Long


    'Error Check ************************************
    Const FunctionName = "MeasureIMulti_HDVIS"
    'Pass
    If CheckPinList(PinList, chHDVIS, FunctionName) = False Then Stop: Exit Sub
    If CheckAvgNum(avgNum, FunctionName) = False Then Stop: Exit Sub
    '************************************************

    'Check Ganged Pins ********************************************
    Call GetActiveChanList(PinList, chHDVIS, Channels)
    Call SeparatePinList(PinList, pinNames)
    If (UBound(pinNames) + 1) * CountActiveSite <> UBound(Channels) + 1 Then
        'Pass
        Call OutputErrMsg(PinList & " Including Ganged Pins. (at MeasureIMulti_HDVIS)")
        Exit Sub
    End If
    '**************************************************************

    'Check MI Mode ************************************************
    Call TheHdw.HDVIS.Pins(PinList).ReadRangesAndMode(CurMode, CurVRange, CurIRange)
    For chan = 0 To UBound(CurMode)
        'Pass
        If CurMode(chan) <> hdvisForceVMeasureI Then
            'Pass
            Call OutputErrMsg(PinList & " Including not MI Mode Pins. (at MeasureIMulti_HDVIS)")
            Exit Sub
        End If
    Next chan
    '**************************************************************

    'Measurement **************************************************
    With TheHdw.HDVIS.Pins(PinList)
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
        'Pass
        For curSite = 0 To CountExistSite
            'Pass
            If IsActiveSite(curSite) Then
                'Pass
                PerPinResults(Pin)(curSite) = Samples(i)
                i = i + 1
            End If
        Next curSite
    Next Pin

    '---------------------------------------------------------

    'Average & Regist ----------------------------------------
    Set m_ResultsI_HDVIS = New Collection
    For Pin = 0 To UBound(pinNames)
        'Pass
        Call m_ResultsI_HDVIS.Add(PerPinResults(Pin), pinNames(Pin))
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

    status = ReadMultiResult(PinName, retResult, m_ResultsV_HDVIS)

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
    
    status = ReadMultiResult(PinName, retResult, m_ResultsI_HDVIS)

    If status = False Then
        Call OutputErrMsg("Not Found Result of " & PinName & " at ReadMIMultiResult()")
        Stop
    End If
    
End Sub

'#Pass-Ok
Private Function GetHdvisFVRange(ByVal ForceV As Double) As HdvisVRange
    
    'Pass
    GetHdvisFVRange = HDVIS10V

End Function

'#Pass-Ok
Private Sub InitMultiDCResult()
'内容:
'   Measure＊Mulitiで使用するコレクションを初期化
'
    Set m_ResultsV_HDVIS = Nothing
    Set m_ResultsI_HDVIS = Nothing

End Sub

'#Pass-Ok
Private Function GetHdvisMIRange(ByVal ClampI As Double) As HdvisIRange

    Dim AbsClampI As Double

    'Pass
    AbsClampI = Abs(ClampI)
    If AbsClampI <= 0.000005 Then       '5uA
        'Pass
        GetHdvisMIRange = hdvis5uA
    ElseIf AbsClampI <= 0.00005 Then    '50uA
        'Pass
        GetHdvisMIRange = hdvis50uA
    ElseIf AbsClampI <= 0.0005 Then     '500uA
        'Pass
        GetHdvisMIRange = hdvis500uA
    ElseIf AbsClampI <= 0.005 Then      '5mA
        'Pass
        GetHdvisMIRange = hdvis5mA
    ElseIf AbsClampI <= 0.05 Then       '50mA
        'Pass
        GetHdvisMIRange = hdvis50mA
    Else
        'Pass
        GetHdvisMIRange = hdvis200mA
    End If

End Function

'#Pass-Ok
Private Function GetHdvisFIRange(ByVal ForceI As Double) As HdvisIRange

    'Pass
    If Abs(ForceI) <= 0.000005 Then     '5uA
        'Pass
        GetHdvisFIRange = hdvis5uA
    ElseIf Abs(ForceI) <= 0.00005 Then  '50uA
        'Pass
        GetHdvisFIRange = hdvis50uA
    ElseIf Abs(ForceI) <= 0.0005 Then   '500uA
        'Pass
        GetHdvisFIRange = hdvis500uA
    ElseIf Abs(ForceI) <= 0.005 Then    '5mA
        'Pass
        GetHdvisFIRange = hdvis5mA
    ElseIf Abs(ForceI) <= 0.05 Then     '50mA
        'Pass
        GetHdvisFIRange = hdvis50mA
    Else
        'Pass
        GetHdvisFIRange = hdvis200mA
    End If

End Function

'#Pass-Ok
Private Function GetHdvisMVRange(ByVal ClampV As Double) As HdvisVRange

    'Pass
    GetHdvisMVRange = HDVIS10V
    
End Function

'#Pass-Ok
Private Function GetHdvisMeasureRange() As Collection
'HDVISの測定レンジを取得する

    Dim pinNameArr() As String
    Dim chListArr() As Long
    Dim tmpHdvisMode() As HdvisMode
    Dim tmpHdvisVrange() As HdvisVRange
    Dim tmpHdvisIrange() As HdvisIRange
    Dim pinLpCnt As Long
    Dim siteLpCnt As Long
    Dim tmpMeasureRng() As Double
    Dim tmpRngVal As Collection
    Dim tmpModeVal As Collection
                                                
    Const FUNCTION_NAME = "GetHdvisMeasureRange"
                                                                                                
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
        TheHdw.HDVIS.Chans(chListArr).ReadRangesAndMode tmpHdvisMode, tmpHdvisVrange, tmpHdvisIrange
                
        'サイトループ
        For siteLpCnt = 0 To CountExistSite
            '現在の設定モードに応じて電圧レンジ値もしくは電流レンジ値を格納する
            Select Case tmpHdvisMode(siteLpCnt)
                Case hdvisForceIMeasureV 'FIMV
                    tmpMeasureRng(siteLpCnt) = HdvisVRangeToValue(tmpHdvisVrange(siteLpCnt))
                
                Case hdvisForceVMeasureI 'FVMI
                    tmpMeasureRng(siteLpCnt) = HdvisIRangeToValue(tmpHdvisIrange(siteLpCnt))
                
                Case hdvisMeasureV       'MV
                    tmpMeasureRng(siteLpCnt) = HdvisVRangeToValue(tmpHdvisVrange(siteLpCnt))
            End Select
        Next siteLpCnt
        
        'ピン名をキーにレンジ値の取得結果をコレクションに登録
        Call tmpRngVal.Add(tmpMeasureRng, pinNameArr(pinLpCnt))
        'ピン名をキーに印加モードの取得結果をコレクションに登録
        Call tmpModeVal.Add(tmpHdvisMode, pinNameArr(pinLpCnt))

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
            Set GetHdvisMeasureRange = Nothing
            Exit Function
        End If
                                                                             
        '★ギャング対象GPのメンバーはすべて同じ状態と信じて
        '  先頭PINのレンジ値と印加モードのみを調べる
        ' （本来はすべてのメンバーの状態から判断を行うのが正しい）
        Dim tmpGangMeasureRng() As Double
        Dim tmpGangHdvisMode() As HdvisMode
        tmpGangMeasureRng = tmpRngVal.Item(pinNameArr(0))
        tmpGangHdvisMode = tmpModeVal.Item(pinNameArr(0))

        '★モード別処理（MIモード時のみ電流レンジ×並列接続数を演算）
        '  (ここでもサイト0の値のみを確認してモードを判断するが
        '  本来はすべてのサイトのモード値を確認するのが正しい
        '  上記のメンバー確認のタイミングですべて確認済みのものを
        '  判定するのが良さそう。)
        Dim tmpGangRngVal As Collection
        Set tmpGangRngVal = New Collection

        If tmpGangHdvisMode(0) <> hdvisForceVMeasureI Then
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
        Set GetHdvisMeasureRange = tmpGangRngVal
                
    'ピンリストの中身が通常ピンのとき(取得結果をそのまま返す)
    Else
        '取得結果を関数の戻り値としてコレクションで返す
        Set GetHdvisMeasureRange = tmpRngVal

    End If
        
End Function

'#Pass-Ok
Private Function GetHdvisForceValue() As Collection
'HDVISのForce値を取得する

    Dim pinNameArr() As String
    Dim chListArr() As Long
    Dim tmpHdvisMode() As HdvisMode
    Dim tmpHdvisVrange() As HdvisVRange
    Dim tmpHdvisIrange() As HdvisIRange
    Dim pinLpCnt As Long
    Dim siteLpCnt As Long
    Dim tmpRetForceVal() As Double
    Dim getForceVal() As Double
    Dim tmpForceVal As Collection
    Dim status As Boolean
                                                                                                
    Const FUNCTION_NAME = "GetHdvisForceValue"
                                                                                                
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
        TheHdw.HDVIS.Chans(chListArr).ReadRangesAndMode tmpHdvisMode, tmpHdvisVrange, tmpHdvisIrange
                
        'サイトループ(CHループ)
        For siteLpCnt = 0 To CountExistSite
            With TheHdw.HDVIS.Chans(chListArr(siteLpCnt))
                '現在の設定モードに応じて電圧レンジ値もしくは電流レンジ値を格納する
                Select Case tmpHdvisMode(siteLpCnt)
                    Case hdvisForceIMeasureV: 'FIMV
                        status = .ReadForceCurrents(tmpHdvisIrange, getForceVal)
                        tmpRetForceVal(siteLpCnt) = getForceVal(0)
                    
                    Case hdvisForceVMeasureI: 'FVMI
                        status = .ReadForceVoltages(tmpHdvisVrange, getForceVal)
                        tmpRetForceVal(siteLpCnt) = getForceVal(0)
                    
                    Case hdvisMeasureV:       'MVモードのときはForce値を0として取り扱う
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
            Set GetHdvisForceValue = Nothing
            Exit Function
        End If
        
        Dim tmpGangForce As Collection
        Set tmpGangForce = New Collection
        Call tmpGangForce.Add(tmpForceVal.Item(pinNameArr(0)), m_PinList)
        'Gang用取得値を取得結果として返す
        Set GetHdvisForceValue = tmpGangForce

    'ピンリストの中身が通常ピンのとき(取得結果をそのまま返す)
    Else
        '取得結果を関数の戻り値としてコレクションで返す
        Set GetHdvisForceValue = tmpForceVal
    End If
            
End Function

'#Pass-Ok
Private Function HdvisIRangeToValue(ByVal Irange As HdvisIRange) As Double
'HDVISのIRange番号をDouble値に変換

    Select Case Irange
        Case hdvis5uA
            HdvisIRangeToValue = 0.000005
        
        Case hdvis50uA
            HdvisIRangeToValue = 0.00005
        
        Case hdvis500uA
            HdvisIRangeToValue = 0.0005
        
        Case hdvis5mA
            HdvisIRangeToValue = 0.005
        
        Case hdvis50mA
            HdvisIRangeToValue = 0.05
        
        Case hdvis200mA
            HdvisIRangeToValue = 0.2
        
        Case Else
            HdvisIRangeToValue = DC_RANGE_UNKNOWN '"Unknown"
    End Select

End Function

'#Pass-Ok
Private Function HdvisVRangeToValue(ByVal Vrange As HdvisVRange) As Double
'HDVISのVRange番号をDouble値に変換

    Select Case Vrange
        Case HDVIS10V
            HdvisVRangeToValue = 10#
        
        Case Else
            HdvisVRangeToValue = DC_RANGE_UNKNOWN
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

    Set IVISVISource_GetForceValue = GetHdvisForceValue
        
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
    
    Set IVISVISource_GetMeasureRange = GetHdvisMeasureRange

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
    
    Call SetMV(m_PinList, site, , ConnectOn)

End Sub
'#V21-Release
Private Sub IVISVISource_SetMVMulti(ByVal ClampV As Double, Optional ByVal ConnectOn As Boolean = True)
    
    Call SetMVMulti(m_PinList, , ConnectOn)

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
    
    status = ReadMultiResult(m_PinList, retResult, m_ResultsV_HDVIS)

    If status = False Then
        Call OutputErrMsg("Not Found Result of " & m_PinList & " at ReadMVMultiResult()")
        Stop
    End If

End Sub

'#Pass-Ok
Private Sub IVISVISource_ReadMIMultiResult(retResult() As Double)
    
    Dim status As Boolean
    
    status = ReadMultiResult(m_PinList, retResult, m_ResultsI_HDVIS)

    If status = False Then
        Call OutputErrMsg("Not Found Result of " & m_PinList & " at ReadMIMultiResult()")
        Stop
    End If

End Sub

'################################################################################

#End If



'#V213-Release 2012.04.27
Private Sub V213_ChangeMIRangeMulti()
' この関数は、Eee-JobCheckPro V1.4 において、Eee-JOBバージョン判定の為のダミー関数です。実行されません。
End Sub

'################################################################################

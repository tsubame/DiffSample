Attribute VB_Name = "XLibSTD_CommonDCMod_V01"
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


Private m_ResultsV_PPMU As Collection
Private m_ResultsV_APMU As Collection
Private m_ResultsI_PPMU As Collection
Private m_ResultsI_APMU As Collection
Private m_ResultsI_DPS As Collection

Private m_FlgDebug As Boolean

'##################################### PPMU #####################################
Public Sub SetFVMI_PPMU( _
    ByVal PinList As String, ByVal ForceV As Variant, ByVal MIRange As PpmuIRange, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal ConnectOn As Boolean = True _
)
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
        Call ConnectPins_PPMU(PinList, site)
    End If

End Sub
'#No-Release
Public Sub SetFVMIMulti_PPMU( _
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

    Dim curSite As Long

    curSite = 0
    'Error Check ************************************
    Const FunctionName = "SetFVMIMulti_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceV, -2 * V, 7 * V, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************

    Call ConvertVariableToArray(ForceV)

            'Main Part ---------------------------------------
            TheHdw.PPMU.Pins(PinList).ForceVoltage(MIRange) = ForceV(curSite)
            '-------------------------------------------------

    If ConnectOn = True Then
        Call ConnectPinsMulti_PPMU(PinList)
    End If

End Sub

Public Sub SetFIMV_PPMU( _
    ByVal PinList As String, ByVal ForceI As Variant, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal FIRange As PpmuIRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
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
        Call ConnectPins_PPMU(PinList, site)
    End If

End Sub
'#No-Release
Public Sub SetFIMVMulti_PPMU( _
    ByVal PinList As String, ByVal ForceI As Variant, _
    Optional ByVal FIRange As PpmuIRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)

'内容:
'    PPMUを電流印加状態に設定する｡(サイト同時)
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

    Dim curSite As Long
    Dim FIRangeBySite() As PpmuIRange

    Dim Channels() As Long

    'Error Check ************************************
    Const FunctionName = "SetFIMVMulti_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceI, -2 * mA, 2 * mA, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************

    Call ConvertVariableToArray(ForceI)

    ReDim FIRangeBySite(1)
    curSite = 0
        If FIRange = DEF_RANGE Then
            FIRangeBySite(curSite) = GetPpmuFIRange(ForceI(curSite))
        Else
            FIRangeBySite(curSite) = FIRange
        End If
   


            'Main Part ---------------------------------------
            TheHdw.PPMU.Pins(PinList).ForceCurrent(FIRangeBySite(curSite)) = ForceI(curSite)
            '-------------------------------------------------

    If ConnectOn = True Then
        Call ConnectPinsMulti_PPMU(PinList)
    End If

End Sub


Public Sub SetMV_PPMU(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE, Optional ConnectOn As Boolean = True)

    Call SetFIMV_PPMU(PinList, 0 * A, site, ppmu200uA, ConnectOn)

End Sub
'#No-Release
Public Sub SetMVMulti_PPMU(ByVal PinList As String, Optional ConnectOn As Boolean = True)

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

    Call SetFIMVMulti_PPMU(PinList, 0 * A, ppmu200uA, ConnectOn)

End Sub


Public Sub DisconnectPins_PPMU(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)

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
'#No-Release
Public Sub DisconnectPinsMulti_PPMU(ByVal PinList As String)

'内容:
'   PPMUをデバイスから切り離す｡(サイト同時)
'
'パラメータ:
'    [PinList]   In   対象ピンリスト。
'
'戻り値:
'
'注意事項:
'    ■PinListに対応するPPMUをデバイスから切り離す｡
'    ■PinListはカンマ区切りのピンリスト､ピングループ指定可能｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■切り離した後、0V印加状態に設定する(接続はしない)。

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

Public Sub ConnectPins_PPMU(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)

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
'#No-Release
Public Sub ConnectPinsMulti_PPMU(ByVal PinList As String)

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


    'Error Check ************************************
    Const FunctionName = "ConnectPinsMulti_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************

            'Main Part ---------------------------------------
            TheHdw.PPMU.Pins(PinList).Connect
            '-------------------------------------------------

End Sub

Public Sub SetGND_PPMU( _
    ByVal PinList As String, Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True _
)

    Call SetFVMI_PPMU(PinList, 0 * V, ppmu2mA, site, ConnectOn)

End Sub

Public Sub ChangeMIRange_PPMU(ByVal PinList As String, ByVal MIRange As PpmuIRange, Optional ByVal site As Long = ALL_SITE)

    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long
    Dim ForceV As Double

    'Error Check ************************************
    Const FunctionName = "ChangeMIRange_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If IsExistSite(site, FunctionName) = False Then Stop: Exit Sub
    '************************************************

    'Site Loop **************************************
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
                        Call DebugMsg("ch" & Channels(chan) & " is not MI Mode.")
                    End If
                End With
            Next chan
            '-------------------------------------------------

        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

End Sub
'#No-Release
Public Sub ChangeMIRangeMulti_PPMU(ByVal PinList As String, ByVal MIRange As PpmuIRange)


'内容:
'   電圧印加状態のPPMUの電流測定レンジを変更する｡(サイト同時)
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

    Dim Channels() As Long
    Dim chan As Long
    Dim ForceV As Double

    'Error Check ************************************
    Const FunctionName = "ChangeMIRangeMulti_PPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************

            'Main Part ---------------------------------------
            Call GetActiveChanList(PinList, chIO, Channels)
            
                With TheHdw.PPMU.Chans(Channels(chan))
                    If .IsForcingVoltage = True Then
                        ForceV = .ForceVoltage(.CurrentRange)
                        .ForceVoltage(MIRange) = ForceV
                    Else
                        Call DebugMsg("ch" & Channels(chan) & " is not MI Mode.")
                    End If
                End With
           
            '-------------------------------------------------

End Sub

Public Sub MeasureV_PPMU( _
    ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE _
)

    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long
    Dim Samples() As Double
    Dim n As Long

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
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then

            'Main Part ---------------------------------------
            Call GetChanList(PinName, curSite, chIO, Channels)

            With TheHdw.PPMU.Chans(Channels(0))
                If .isForcingCurrent = False Then
                    Call DebugMsg(PinName & " is not MV Mode (at MeasureV_PPMU")
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

End Sub

Public Sub MeasureI_PPMU( _
    ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE _
)

    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long
    Dim Samples() As Double
    Dim n As Long

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
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then

            'Main Part ---------------------------------------
            Call GetChanList(PinName, curSite, chIO, Channels)

            With TheHdw.PPMU.Chans(Channels(0))
                If .IsForcingVoltage = False Then
                    Call DebugMsg(PinName & " is not MI Mode (at MeasureI_PPMU")
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

End Sub

Public Sub MeasureVMulti_PPMU(ByVal PinList As String, ByVal avgNum As Long)

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
    If i <> UBound(Samples) + 1 Then
        Call DebugMsg("Error Happened. (at MeasureVMulti_PPMU)")
        Exit Sub
    End If
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

Public Sub MeasureIMulti_PPMU(ByVal PinList As String, ByVal avgNum As Long)

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
    If i <> UBound(Samples) + 1 Then
        Call DebugMsg("Error Happened. (at MeasureIMulti_PPMU)")
        Exit Sub
    End If
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

Private Function GetPpmuFIRange(ByVal ForceI As Double) As PpmuIRange

    Dim AbsForceI As Double

    AbsForceI = Abs(ForceI)
    If AbsForceI <= 0.0002 Then         '200uA
        GetPpmuFIRange = ppmu200uA
    Else
        GetPpmuFIRange = ppmu2mA
    End If

End Function

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
'################################################################################



'##################################### BPMU #####################################
Public Sub SetFVMI_BPMU( _
    ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, _
    Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal FVRange As BpmuVRange = DEF_RANGE, Optional ByVal MIRange As BpmuIRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
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
        Call ConnectPins_BPMU(PinList, site)
    End If

End Sub
'#No-Release
Public Sub SetFVMIMulti_BPMU( _
    ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, _
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
    Dim curSite As Long
    Dim FVRangeBySite() As BpmuVRange

    'Error Check ************************************
    Const FunctionName = "SetFVMIMulti_BPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceV, -24 * V, 24 * V, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 0 * mA, 200 * mA, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************

    Call ConvertVariableToArray(ForceV)

    ReDim FVRangeBySite(1)
    curSite = 0
        If FVRange = DEF_RANGE Then
            FVRangeBySite(curSite) = GetBpmuFVRange(ForceV(curSite))
        Else
            FVRangeBySite(curSite) = FVRange
        End If


    If MIRange = DEF_RANGE Then
        MIRange = GetBpmuMIRange(ClampI)
    End If

            'Main Part ---------------------------------------
            With TheHdw.BPMU.Pins(PinList)
                .ClampCurrent(MIRange) = ClampI
                .ForceVoltage(FVRangeBySite(curSite)) = ForceV(curSite)
                Call .ModeFVMI(MIRange, FVRangeBySite(curSite))
            End With
            '-------------------------------------------------

    If ConnectOn = True Then
        Call ConnectPinsMulti_BPMU(PinList)
    End If

End Sub

Public Sub SetFIMV_BPMU( _
    ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, _
    Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal FIRange As BpmuIRange = DEF_RANGE, Optional ByVal MVRange As BpmuVRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
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
        Call ConnectPins_BPMU(PinList, site)
    End If

End Sub
'#No-Release
Public Sub SetFIMVMulti_BPMU( _
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

    Dim curSite As Long
    Dim FIRangeBySite() As BpmuIRange

    'Error Check ************************************
    Const FunctionName = "SetFIMVMulti_BPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceI, -200 * mA, 200 * mA, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampV, 0 * V, 24 * V, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************

    Call ConvertVariableToArray(ForceI)

    ReDim FIRangeBySite(1)
    curSite = 0
        If FIRange = DEF_RANGE Then
            FIRangeBySite(curSite) = GetBpmuFIRange(ForceI(curSite))
        Else
            FIRangeBySite(curSite) = FIRange
        End If

    If MVRange = DEF_RANGE Then
        MVRange = GetBpmuMVRange(ClampV)
    End If

            'Main Part ---------------------------------------
            With TheHdw.BPMU.Pins(PinList)
                .ClampVoltage(MVRange) = ClampV
                .ForceCurrent(FIRangeBySite(curSite)) = ForceI(curSite)
                Call .ModeFIMV(FIRangeBySite(curSite), MVRange)
            End With
            '------------------------------------------------

    If ConnectOn = True Then
        Call ConnectPinsMulti_BPMU(PinList)
    End If

End Sub

Public Sub SetMV_BPMU( _
    ByVal PinList As String, ByVal ClampV As Double, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal MVRange As BpmuVRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)

    Call SetFIMV_BPMU(PinList, 0 * A, ClampV, site, , MVRange, ConnectOn)

End Sub
'#No-Release
Public Sub SetMVMulti_BPMU( _
    ByVal PinList As String, ByVal ClampV As Double, _
    Optional ByVal MVRange As BpmuVRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)

    '内容:
'   BPMUを0A印加状態に設定する｡
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

    Call SetFIMVMulti_BPMU(PinList, 0 * A, ClampV, , MVRange, ConnectOn)

End Sub

Public Sub DisconnectPins_BPMU(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)

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
'#No-Release
Public Sub DisconnectPinsMulti_BPMU(ByVal PinList As String)

    Dim curSite As Long

    'Error Check ************************************
    Const FunctionName = "DisconnectPinsMulti_BPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************


            'Main Part ---------------------------------------
            TheHdw.Digital.relays.Pins(PinList).DisconnectPins
            '-------------------------------------------------

End Sub
Public Sub ConnectPins_BPMU(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)

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
'#No-Release
Public Sub ConnectPinsMulti_BPMU(ByVal PinList As String)



    'Error Check ************************************
    Const FunctionName = "DisconnectPinsMulti_BPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************



            'Main Part ---------------------------------------
            TheHdw.BPMU.Pins(PinList).Connect
            '-------------------------------------------------


End Sub
Public Sub SetGND_BPMU( _
    ByVal PinList As String, Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True _
)

    Call SetFVMI_BPMU(PinList, 0 * V, 200 * mA, site, , , ConnectOn)

End Sub

Public Sub ChangeMIRange_BPMU( _
    ByVal PinList As String, ByVal ClampI As Double, Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal MIRange As BpmuIRange = DEF_RANGE _
)
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long
    Dim ForceV As Double
    Dim CurIRanges() As Long
    Dim CurVRanges() As Long

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
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then

            'Main Part ---------------------------------------
            If TheHdw.BPMU.Pins(PinList).BpmuIsMeasuringCurrent(curSite) = False Then
                Call DebugMsg(PinList & " is not MI Mode. (at ChangeMIRange_BPMU)")
            Else
                Call GetChanList(PinList, curSite, chIO, Channels)
                For chan = 0 To UBound(Channels)
                    With TheHdw.BPMU.Chans(Channels(chan))
                        Call .ReadRanges(CurIRanges, CurVRanges)
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

End Sub

'#No-Release
Public Sub ChangeMIRangeMulti_BPMU( _
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



    Dim Channels() As Long
    Dim ForceV As Double
    Dim CurIRanges() As Long
    Dim CurVRanges() As Long

    'Error Check ************************************
    Const FunctionName = "ChangeMIRangeMulti_BPMU"
    If CheckPinList(PinList, chIO, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 0, 200 * mA, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************

    If MIRange = DEF_RANGE Then
        MIRange = GetBpmuMIRange(ClampI)
    End If

'再考の余地有り

            'Main Part ---------------------------------------
                    With TheHdw.BPMU.Pins(PinList)
                        Call .ReadRanges(CurIRanges, CurVRanges)
                        .ClampCurrent(MIRange) = ClampI
                        Call .ModeFVMI(MIRange, CurVRanges(0))
                    End With
                    
'                   Chans指定
'                Call GetActiveChanList(PinList, chIO, Channels)
'                    With TheHdw.BPMU.Chans(Channels)
'                        Call .ReadDriverRanges(CurIRanges, CurVRanges)
'                        .ClampCurrent(MIRange) = ClampI
'                        Call .ModeFVMI(MIRange, CurVRanges(0))
'                    End With

'            End If
            '-------------------------------------------------


End Sub

Public Sub MeasureV_BPMU( _
    ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE _
)
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long
    Dim Samples() As Double
    Dim n As Long

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
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then

            'Main Part ---------------------------------------
            Call GetChanList(PinName, curSite, chIO, Channels)
            With TheHdw.BPMU
                If .Pins(PinName).BpmuIsMeasuringVoltage(curSite) = False Then
                    Call DebugMsg(PinName & " is not MV Mode (at MeasureV_BPMU")
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

End Sub

Public Sub MeasureI_BPMU( _
    ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE _
)
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long
    Dim Samples() As Double
    Dim n As Long

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
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then

            'Main Part ---------------------------------------
            Call GetChanList(PinName, curSite, chIO, Channels)
            With TheHdw.BPMU
                If .Pins(PinName).BpmuIsMeasuringCurrent(curSite) = False Then
                    Call DebugMsg(PinName & " is not MV Mode (at MeasureI_BPMU")
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

End Sub

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

Private Function GetBpmuMIRange(ByVal ClampI As Double) As BpmuIRange

    Dim AbsClampI As Double

    AbsClampI = Abs(ClampI)
    If AbsClampI <= 0.000002 Then       '2uA
        GetBpmuMIRange = bpmu2uA
    ElseIf AbsClampI <= 0.00002 Then    '20uA
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
'################################################################################

'##################################### APMU #####################################
Public Sub SetFVMI_APMU( _
    ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, _
    Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal FVRange As ApmuVRange = DEF_RANGE, Optional ByVal MIRange As ApmuIRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
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
        Call ConnectPins_APMU(PinList, site)
    End If

End Sub

'################################################################################
'##################################### APMU #####################################
'##################################### For PowerDown ############################
Public Sub SetFVMI_APMUoff( _
    ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, _
    Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal FVRange As ApmuVRange = DEF_RANGE, Optional ByVal MIRange As ApmuIRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
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

'### For Power Down ###
    MIRange = 50 * mA
    ClampI = 5 * mA
'######################

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
        Call ConnectPins_APMU(PinList, site)
    End If

End Sub

'#No-Release
Public Sub SetFVMIMulti_APMU( _
    ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, _
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

    Dim curSite As Long
    Dim FVRangeBySite() As ApmuVRange

    'Error Check ************************************
    Const FunctionName = "SetFVMIMulti_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceV, -35 * V, 35 * V, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 0 * mA, 50 * mA * 8, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************

    Call ConvertVariableToArray(ForceV)

    ReDim FVRangeBySite(1)
    curSite = 0
        If FVRange = DEF_RANGE Then
            FVRangeBySite(curSite) = GetApmuFVRange(ForceV(curSite))
        Else
            FVRangeBySite(curSite) = FVRange
        End If

    If MIRange = DEF_RANGE Then
        MIRange = GetApmuMIRange(ClampI)
    End If

            'Main Part ---------------------------------------
            With TheHdw.APMU.Pins(PinList)
                .alarm = False
                .ClampCurrent(MIRange) = ClampI
                .ForceVoltage(FVRangeBySite(curSite)) = ForceV(curSite)
                Call .ModeFVMI(MIRange)
                .Gate = True
            End With
            '-------------------------------------------------

    If ConnectOn = True Then
        Call ConnectPinsMulti_APMU(PinList)
    End If

End Sub
Public Sub SetFIMV_APMU( _
    ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, _
    Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal FIRange As ApmuIRange = DEF_RANGE, Optional ByVal MVRange As ApmuVRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
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
        Call ConnectPins_APMU(PinList, site)
    End If

End Sub

'#No-Release
Public Sub SetFIMVMulti_APMU( _
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
'    ■全サイト同じ値を設定。
'    ■ClampVでクランプ電圧を設定｡
'    ■非アクティブサイトに対しては何もしない｡
'    ■FIRangeで印加電流レンジを設定。-1(デフォルト)の場合、ForceIからレンジを判定して設定する。
'    ■MVRangeで測定レンジを指定。-1(デフォルト)の場合、ClampVからレンジを判定して設定する。
'    ■ConnectOnをTrue(デフォルト)にすると、設定とコネクトを一度にする。
'    ■ConnectOnをFalseにすると､設定のみでコネクトは行なわない｡ (既にコネクトされている場合はそのまま)

    Dim curSite As Long
    Dim FIRangeBySite() As ApmuIRange

    'Error Check ************************************
    Const FunctionName = "SetFIMVMulti_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceI, -50 * mA * 8, 50 * mA * 8, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampV, 0 * V, 35 * V, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************

    Call ConvertVariableToArray(ForceI)

    ReDim FIRangeBySite(1)
    curSite = 0
        If FIRange = DEF_RANGE Then
            FIRangeBySite(curSite) = GetApmuFIRange(ForceI(curSite))
        Else
            FIRangeBySite(curSite) = FIRange
        End If

    If MVRange = DEF_RANGE Then
        MVRange = GetApmuMVRange(ClampV)
    End If


            'Main Part ---------------------------------------
            With TheHdw.APMU.Pins(PinList)
                .alarm = False
                .ClampVoltage(MVRange) = ClampV
                .ForceCurrent(FIRangeBySite(curSite)) = ForceI(curSite)
                Call .ModeFIMV(MVRange)
                .Gate = True
            End With
            '-------------------------------------------------


    If ConnectOn = True Then
        Call ConnectPinsMulti_APMU(PinList)
    End If

End Sub
Public Sub SetMV_APMU( _
    ByVal PinList As String, ByVal ClampV As Double, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal MVRange As ApmuVRange = DEF_RANGE, _
    Optional ByVal ConnectOn As Boolean = True _
)
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
        Call ConnectPins_APMU(PinList, site)
    End If

End Sub

'No-Release
Public Sub SetMVMulti_APMU( _
    ByVal PinList As String, ByVal ClampV As Double, _
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

    Dim curSite As Long


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
            With TheHdw.APMU.Pins(PinList)
                .alarm = False
                .ClampVoltage(MVRange) = ClampV
                Call .ModeMV(MVRange)
                .Gate = True
            End With
            '-------------------------------------------------

    If ConnectOn = True Then
        Call ConnectPinsMulti_APMU(PinList)
    End If

End Sub

Public Sub DisconnectPins_APMU(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)

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
'#No-Release
Public Sub DisconnectPinsMulti_APMU(ByVal PinList As String)

    Dim curSite As Long

    'Error Check ************************************
    Const FunctionName = "DisconnectPinsMulti_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************


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


End Sub

Public Sub ConnectPins_APMU(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)

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

'#No-Release
Public Sub ConnectPinsMulti_APMU(ByVal PinList As String)

    Dim curSite As Long

    'Error Check ************************************
    Const FunctionName = "ConnectPinsMulti_APMU"
    If CheckPinList(PinList, chAPMU, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************


            'Main Part ---------------------------------------
            TheHdw.APMU.Pins(PinList).relay = True
            '-------------------------------------------------


End Sub
Public Sub SetGND_APMU( _
    ByVal PinList As String, Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True _
)

    Call SetFVMI_APMU(PinList, 0 * V, 50 * mA, site, , , ConnectOn)

End Sub

Public Sub ChangeMIRange_APMU( _
    ByVal PinList As String, ByVal ClampI As Double, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal MIRange As ApmuIRange = DEF_RANGE _
)
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long

    Dim CurMode() As ApmuMode
    Dim CurVRange() As ApmuVRange
    Dim CurIRange() As ApmuIRange

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
                        Call DebugMsg("APMU(" & Channels(chan) & ") is not MI Mode. (at ChangeMIRange_APMU)")
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

End Sub

'No-Release

Public Sub ChangeMIRangeMulti_APMU( _
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

    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long

    Dim CurMode() As ApmuMode
    Dim CurVRange() As ApmuVRange
    Dim CurIRange() As ApmuIRange

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
            Call GetActiveChanList(PinList, chAPMU, Channels)
                With TheHdw.APMU.Chans(Channels)
                    Call .ReadRangesAndMode(CurMode, CurVRange, CurIRange)
                    If CurMode(0) <> apmuForceVMeasureI Then
                        Call DebugMsg("APMU(" & Channels(chan) & ") is not MI Mode. (at ChangeMIRange_APMU)")
                    Else
                        .ClampCurrent(MIRange) = ClampI
                        Call .ModeFVMI(MIRange)
                    End If
                End With
            '-------------------------------------------------


End Sub


Public Sub MeasureV_APMU( _
    ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal UseLPF As Boolean = False _
)
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long
    Dim Samples() As Double
    Dim n As Long

    Dim CurMode() As ApmuMode
    Dim CurVRange() As ApmuVRange
    Dim CurIRange() As ApmuIRange

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
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then

            'Main Part ---------------------------------------
            Call GetChanList(PinName, curSite, chAPMU, Channels)

            With TheHdw.APMU.Pins(PinName)
                Call .ReadRangesAndMode(CurMode, CurVRange, CurIRange)
                If CurMode(0) = apmuForceVMeasureI Then
                    Call DebugMsg(PinName & " is not MV Mode (at MeasureV_APMU")
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

End Sub

Public Sub MeasureI_APMU( _
    ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal UseLPF As Boolean = False _
)
    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long
    Dim Samples() As Double
    Dim n As Long

    Dim CurMode() As ApmuMode
    Dim CurVRange() As ApmuVRange
    Dim CurIRange() As ApmuIRange

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
    siteStatus = TheExec.sites.SelectFirst
    Do While siteStatus <> loopDone
        curSite = TheExec.sites.SelectedSite
        If curSite = site Or site = ALL_SITE Then

            'Main Part ---------------------------------------
            Call GetChanList(PinName, curSite, chAPMU, Channels)
            If UBound(Channels) >= 1 Then
                Call DebugMsg("Don't Support Multi Pins. (at MeasureI_APMU)")
            Else
                With TheHdw.APMU.Pins(PinName)
                    Call .ReadRangesAndMode(CurMode, CurVRange, CurIRange)
                    If CurMode(0) <> apmuForceVMeasureI Then
                        Call DebugMsg(PinName & " is not MI Mode (at MeasureI_APMU")
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
            End If
            '-------------------------------------------------

        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

End Sub

Public Sub MeasureVMulti_APMU(ByVal PinList As String, ByVal avgNum As Long, Optional ByVal UseLPF As Boolean = False)

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
        Call DebugMsg(PinList & " Including Ganged Pins. (at MeasureVMulti_APMU)")
        Exit Sub
    End If
    '**************************************************************

    'Check MV Mode ************************************************
    Call TheHdw.APMU.Pins(PinList).ReadRangesAndMode(CurMode, CurVRange, CurIRange)
    For chan = 0 To UBound(CurMode)
        If CurMode(chan) = apmuForceVMeasureI Then
            Call DebugMsg(PinList & " Including not MV Mode Pins. (at MeasureVMulti_APMU)")
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

    If i <> UBound(Samples) + 1 Then
        Call DebugMsg("Error Happened. (at MeasureVMulti_APMU)")
        Exit Sub
    End If
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

Public Sub MeasureIMulti_APMU(ByVal PinList As String, ByVal avgNum As Long, Optional ByVal UseLPF As Boolean = False)

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
        Call DebugMsg(PinList & " Including Ganged Pins. (at MeasureIMulti_APMU)")
        Exit Sub
    End If
    '**************************************************************

    'Check MI Mode ************************************************
    Call TheHdw.APMU.Pins(PinList).ReadRangesAndMode(CurMode, CurVRange, CurIRange)
    For chan = 0 To UBound(CurMode)
        If CurMode(chan) <> apmuForceVMeasureI Then
            Call DebugMsg(PinList & " Including not MI Mode Pins. (at MeasureIMulti_APMU)")
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

    If i <> UBound(Samples) + 1 Then
        Call DebugMsg("Error Happened. (at MeasureIMulti_APMU)")
        Exit Sub
    End If
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
'################################################################################

'##################################### DPS ######################################
Public Sub SetFVMI_DPS( _
    ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, _
    ByVal MIRange As DpsIRange, Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True _
)
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
        Call ConnectPins_DPS(PinList, site)
    End If

End Sub

'No-Release
Public Sub SetFVMIMulti_DPS( _
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

    Dim curSite As Long
    Dim Channels() As Long

    'Error Check ************************************
    Const FunctionName = "SetFVMIMulti_DPS"
    If CheckPinList(PinList, chDPS, FunctionName) = False Then Stop: Exit Sub
    If CheckForceVariantValue(ForceV, 0 * V, 10 * V, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 50 * mA, 1 * A, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************

    Call ConvertVariableToArray(ForceV)

    curSite = 0

            'Main Part ---------------------------------------
            With TheHdw.DPS.Pins(PinList)
                .CurrentRange = MIRange
                .CurrentLimit = ClampI
                .forceValue(dpsPrimaryVoltage) = ForceV(curSite)
            End With
            '-------------------------------------------------

    If ConnectOn = True Then
        Call ConnectPinsMulti_DPS(PinList)
    End If

End Sub

Public Sub SetGND_DPS( _
    ByVal PinList As String, Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True _
)

    Call SetFVMI_DPS(PinList, 0 * V, 1 * A, dps1a, site, ConnectOn)

End Sub

Public Sub DisconnectPins_DPS(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)

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

    Call DisconnectGangedPins_DPS(PinList, site)

End Sub
'#No-Release
Public Sub DisconnectPinsMulti_DPS(ByVal PinList As String)

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

    Call DisconnectGangedPins_DPS(PinList, ALL_SITE)

End Sub


Private Sub DisconnectGangedPins_DPS(ByVal PinList As String, ByVal site As Long)

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

Public Sub ConnectPins_DPS(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)

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
'#No-Release
Public Sub ConnectPinsMulti_DPS(ByVal PinList As String)

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

Public Sub ChangeMIRange_DPS( _
    ByVal PinList As String, ByVal ClampI As Double, _
    ByVal MIRange As DpsIRange, Optional ByVal site As Long = ALL_SITE _
)
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

'No-Release
Public Sub ChangeMIRangeMulti_DPS( _
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

    Dim Channels() As Long

    'Error Check ************************************
    Const FunctionName = "ChangeMIRangeMulti_DPS"
    If CheckPinList(PinList, chDPS, FunctionName) = False Then Stop: Exit Sub
    If CheckClampValue(ClampI, 50 * mA, 1 * A, FunctionName) = False Then Stop: Exit Sub
'    If IsExistSite(Site, functionName) = False Then Stop: Exit Sub
    '************************************************

' 要検討
                       'Main Part ---------------------------------------
'Pins指定
                With TheHdw.DPS.Pins(PinList)
                    .CurrentRange = MIRange
                    .CurrentLimit = ClampI
                End With
                       
'Chans指定
'            Call GetActiveChanList(PinList, chDPS, Channels)
'                With TheHdw.DPS.Chans(Channels)
'                    .CurrentRange = MIRange
'                    .CurrentLimit = ClampI
'                End With

            '-------------------------------------------------

End Sub
Public Sub MeasureI_DPS(ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, Optional ByVal site As Long = ALL_SITE)

    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim Channels() As Long
    Dim chan As Long
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
            If UBound(Channels) >= 1 Then
                Call DebugMsg("Don't Support Multi Pins. (at MeasureI_DPS)")
            Else
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
            End If
            '-------------------------------------------------

        End If
        siteStatus = TheExec.sites.SelectNext(siteStatus)
    Loop
    '************************************************

End Sub

Public Sub MeasureIMulti_DPS(ByVal PinList As String, ByVal avgNum As Long)

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
        Call DebugMsg(PinList & " Including Ganged Pins. (at MeasureIMulti_DPS)")
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

    If i <> UBound(Samples) + 1 Then
        Call DebugMsg("Error Happened. (at MeasureIMulti_DPS)")
        Exit Sub
    End If
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
'################################################################################

'##################################### ALL ######################################
Public Sub SetFVMI( _
    ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal ConnectOn As Boolean = True _
)

    Dim chanType As chtype
    Dim DpsMIRange As DpsIRange

    chanType = GetChanType(PinList)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        Call SetFVMI_PPMU(PinList, ForceV, GetPpmuMIRange(ClampI), site, ConnectOn)

    Case chAPMU '---------------------------------------------------------
        Call SetFVMI_APMU(PinList, ForceV, Abs(ClampI), site, , , ConnectOn)

    Case chDPS '----------------------------------------------------------
        Call SetFVMI_DPS(PinList, ForceV, CreateLimit(Abs(ClampI), 50 * mA, 1 * A), GetDpsMIRange(ClampI), site, ConnectOn)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinList & " is Invalid Pin List at SetFVMI()")
        Stop
    End Select

End Sub

'#No-Release
Public Sub SetFVMIMulti( _
    ByVal PinList As String, ByVal ForceV As Variant, ByVal ClampI As Double, _
     Optional ByVal ConnectOn As Boolean = True _
)

    Dim chanType As chtype
    Dim DpsMIRange As DpsIRange

    chanType = GetChanType(PinList)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        Call SetFVMIMulti_PPMU(PinList, ForceV, GetPpmuMIRange(ClampI), ConnectOn)

    Case chAPMU '---------------------------------------------------------
        Call SetFVMIMulti_APMU(PinList, ForceV, Abs(ClampI), , , ConnectOn)

    Case chDPS '----------------------------------------------------------
        Call SetFVMIMulti_DPS(PinList, ForceV, CreateLimit(Abs(ClampI), 50 * mA, 1 * A), GetDpsMIRange(ClampI), ConnectOn)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinList & " is Invalid Pin List at SetFVMIMulti()")
        Stop
    End Select

End Sub

Public Sub SetFIMV( _
    ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, _
    Optional ByVal site As Long = ALL_SITE, Optional ByVal ConnectOn As Boolean = True _
)
    Dim chanType As chtype

    chanType = GetChanType(PinList)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        Call SetFIMV_PPMU(PinList, ForceI, site, , ConnectOn)

    Case chAPMU '---------------------------------------------------------
        Call SetFIMV_APMU(PinList, ForceI, Abs(ClampV), site, , , ConnectOn)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinList & " is Invalid Pin List at SetFIMV()")
        Stop
    End Select

End Sub

'#No-Release
Public Sub SetFIMVMulti( _
    ByVal PinList As String, ByVal ForceI As Variant, ByVal ClampV As Double, _
    Optional ByVal ConnectOn As Boolean = True _
)
    Dim chanType As chtype

    chanType = GetChanType(PinList)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        Call SetFIMVMulti_PPMU(PinList, ForceI, , ConnectOn)

    Case chAPMU '---------------------------------------------------------
        Call SetFIMVMulti_APMU(PinList, ForceI, Abs(ClampV), , , ConnectOn)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinList & " is Invalid Pin List at SetFIMVMulti()")
        Stop
    End Select

End Sub

Public Sub SetMV( _
    ByVal PinList As String, ByVal ClampV As Double, Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True _
)

    Dim chanType As chtype

    chanType = GetChanType(PinList)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        Call SetMV_PPMU(PinList, site, ConnectOn)

    Case chAPMU '---------------------------------------------------------
        Call SetMV_APMU(PinList, ClampV, site, , ConnectOn)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinList & " is Invalid Pin List at SetMV()")
        Stop
    End Select

End Sub

'#No-Release
Public Sub SetMVMulti( _
    ByVal PinList As String, ByVal ClampV As Double, _
    Optional ByVal ConnectOn As Boolean = True _
)

    Dim chanType As chtype

    chanType = GetChanType(PinList)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        Call SetMVMulti_PPMU(PinList, ConnectOn)

    Case chAPMU '---------------------------------------------------------
        Call SetMVMulti_APMU(PinList, ClampV, , ConnectOn)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinList & " is Invalid Pin List at SetMVMulti()")
        Stop
    End Select

End Sub



Public Sub ConnectPins(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)

    Dim chanType As chtype

    chanType = GetChanType(PinList)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        Call ConnectPins_PPMU(PinList, site)

    Case chAPMU '---------------------------------------------------------
        Call ConnectPins_APMU(PinList, site)

    Case chDPS '----------------------------------------------------------
        Call ConnectPins_DPS(PinList, site)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinList & " is Invalid Pin List at ConnectPins()")
        Stop
    End Select

End Sub

'No-Release
Public Sub ConnectPinsMulti(ByVal PinList As String)

    Dim chanType As chtype

    chanType = GetChanType(PinList)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        Call ConnectPinsMulti_PPMU(PinList)

    Case chAPMU '---------------------------------------------------------
        Call ConnectPinsMulti_APMU(PinList)

    Case chDPS '----------------------------------------------------------
        Call ConnectPinsMulti_DPS(PinList)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinList & " is Invalid Pin List at ConnectPinsMulti()")
        Stop
    End Select

End Sub

Public Sub DisconnectPins(ByVal PinList As String, Optional ByVal site As Long = ALL_SITE)

    Dim chanType As chtype

    chanType = GetChanType(PinList)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        Call DisconnectPins_PPMU(PinList, site)

    Case chAPMU '---------------------------------------------------------
        Call DisconnectPins_APMU(PinList, site)

    Case chDPS '----------------------------------------------------------
        Call DisconnectPins_DPS(PinList, site)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinList & " is Invalid Pin List at DisconnectPins()")
        Stop
    End Select

End Sub

'No-Release
Public Sub DisconnectPinsMulti(ByVal PinList As String)

    Dim chanType As chtype

    chanType = GetChanType(PinList)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        Call DisconnectPinsMulti_PPMU(PinList)

    Case chAPMU '---------------------------------------------------------
        Call DisconnectPinsMulti_APMU(PinList)

    Case chDPS '----------------------------------------------------------
        Call DisconnectPinsMulti_DPS(PinList)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinList & " is Invalid Pin List at DisconnectPinsMulti()")
        Stop
    End Select

End Sub


Public Sub SetGND( _
    ByVal PinList As String, Optional ByVal site As Long = ALL_SITE, _
    Optional ByVal ConnectOn As Boolean = True _
)

    Dim chanType As chtype

    chanType = GetChanType(PinList)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        Call SetGND_PPMU(PinList, site, ConnectOn)

    Case chAPMU '---------------------------------------------------------
        Call SetGND_APMU(PinList, site, ConnectOn)

    Case chDPS '----------------------------------------------------------
        Call SetGND_DPS(PinList, site, ConnectOn)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinList & " is Invalid Pin List at SetGND()")
        Stop
    End Select

End Sub

Public Sub ChangeMIRange(ByVal PinList As String, ByVal ClampI As Double, Optional ByVal site As Long = ALL_SITE)

    Dim chanType As chtype

    chanType = GetChanType(PinList)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        Call ChangeMIRange_PPMU(PinList, GetPpmuMIRange(ClampI), site)

    Case chAPMU '---------------------------------------------------------
        Call ChangeMIRange_APMU(PinList, Abs(ClampI), site)

    Case chDPS '----------------------------------------------------------
        Call ChangeMIRange_DPS(PinList, CreateLimit(Abs(ClampI), 50 * mA, 1 * A), GetDpsMIRange(ClampI), site)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinList & " is Invalid Pin List at ChangeMIRange()")
        Stop
    End Select

End Sub

'No-Release
Public Sub ChangeMIRangeMulti(ByVal PinList As String, ByVal ClampI As Double)

    Dim chanType As chtype

    chanType = GetChanType(PinList)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        Call ChangeMIRangeMulti_PPMU(PinList, GetPpmuMIRange(ClampI))

    Case chAPMU '---------------------------------------------------------
        Call ChangeMIRangeMulti_APMU(PinList, Abs(ClampI))

    Case chDPS '----------------------------------------------------------
        Call ChangeMIRangeMulti_DPS(PinList, CreateLimit(Abs(ClampI), 50 * mA, 1 * A), GetDpsMIRange(ClampI))

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinList & " is Invalid Pin List at ChangeMIRangeMulti()")
        Stop
    End Select

End Sub



Public Sub MeasureV(ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, Optional ByVal site As Long = ALL_SITE)

    Dim chanType As chtype

    chanType = GetChanType(PinName)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        Call MeasureV_PPMU(PinName, retResult, avgNum, site)

    Case chAPMU '---------------------------------------------------------
        Call MeasureV_APMU(PinName, retResult, avgNum, site)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinName & " is Invalid Pin Name at MeasureV()")
        Stop
    End Select

End Sub

Public Sub MeasureI(ByVal PinName As String, ByRef retResult() As Double, ByVal avgNum As Long, Optional ByVal site As Long = ALL_SITE)

    Dim chanType As chtype

    chanType = GetChanType(PinName)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        Call MeasureI_PPMU(PinName, retResult, avgNum, site)

    Case chAPMU '---------------------------------------------------------
        Call MeasureI_APMU(PinName, retResult, avgNum, site)

    Case chDPS '----------------------------------------------------------
        Call MeasureI_DPS(PinName, retResult, avgNum, site)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinName & " is Invalid Pin Name at MeasureI()")
        Stop
    End Select

End Sub

Public Sub MeasureIMulti(ByVal PinList As String, ByVal avgNum As Long)

    Dim chanType As chtype

    Call InitMultiDCResult

    chanType = GetChanType(PinList)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        Call MeasureIMulti_PPMU(PinList, avgNum)

    Case chAPMU '---------------------------------------------------------
        Call MeasureIMulti_APMU(PinList, avgNum)

    Case chDPS '----------------------------------------------------------
        Call MeasureIMulti_DPS(PinList, avgNum)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinList & " is Invalid Pin List at MeasureIMulti()")
        Stop
    End Select

End Sub

Public Sub MeasureVMulti(ByVal PinList As String, ByVal avgNum As Long)

    Dim chanType As chtype

    Call InitMultiDCResult

    chanType = GetChanType(PinList)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        Call MeasureVMulti_PPMU(PinList, avgNum)

    Case chAPMU '---------------------------------------------------------
        Call MeasureVMulti_APMU(PinList, avgNum)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinList & " is Invalid Pin List at MeasureVMulti()")
        Stop
    End Select

End Sub

Public Sub ReadMVMultiResult(ByVal PinName As String, ByRef retResult() As Double)

    Dim chanType As chtype
    Dim status As Boolean

    chanType = TheExec.DataManager.chanType(PinName)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        status = ReadMultiResult(PinName, retResult, m_ResultsV_PPMU)

    Case chAPMU '---------------------------------------------------------
        status = ReadMultiResult(PinName, retResult, m_ResultsV_APMU)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinName & " is Invalid Pin Name at ReadMVMultiResult()")
        Stop
        Exit Sub
    End Select

    If status = False Then
        Call DebugMsg("Not Found Result of " & PinName & " at ReadMVMultiResult()")
        Stop
    End If

End Sub

Public Sub ReadMIMultiResult(ByVal PinName As String, ByRef retResult() As Double)

    Dim chanType As chtype
    Dim status As Boolean

    chanType = TheExec.DataManager.chanType(PinName)

    Select Case chanType
    Case chIO '-----------------------------------------------------------
        status = ReadMultiResult(PinName, retResult, m_ResultsI_PPMU)

    Case chAPMU '---------------------------------------------------------
        status = ReadMultiResult(PinName, retResult, m_ResultsI_APMU)

    Case chDPS '----------------------------------------------------------
        status = ReadMultiResult(PinName, retResult, m_ResultsI_DPS)

    Case Else '-----------------------------------------------------------
        Call DebugMsg(PinName & " is Invalid Pin Name at ReadMIMultiResult()")
        Stop
        Exit Sub
    End Select

    If status = False Then
        Call DebugMsg("Not Found Result of " & PinName & " at ReadMIMultiResult()")
        Stop
    End If

End Sub

Private Function ReadMultiResult(ByVal PinName As String, ByRef retResult() As Double, ByRef Results As Collection) As Boolean

    Dim site As Long
    Dim result As Variant

    On Error GoTo NOT_FOUND
    result = Results(PinName)
    On Error GoTo 0

    For site = 0 To CountExistSite
        retResult(site) = result(site)
    Next site

    ReadMultiResult = True
    Exit Function

NOT_FOUND:
    ReadMultiResult = False

End Function

Public Sub InitMultiDCResult()

    Set m_ResultsV_PPMU = Nothing
    Set m_ResultsV_APMU = Nothing
    Set m_ResultsI_PPMU = Nothing
    Set m_ResultsI_APMU = Nothing
    Set m_ResultsI_DPS = Nothing

End Sub
'################################################################################

'#################################### COMMON ####################################
Private Function ConvertVariableToArray(ByRef DstVar As Variant) As Boolean

    Dim VarArray() As Double
    Dim site As Long

    If IsArray(DstVar) Then
        If UBound(DstVar) <> CountExistSite Then
            ConvertVariableToArray = False
        Else
            ConvertVariableToArray = True
        End If
    Else
        ReDim VarArray(CountExistSite)

        For site = 0 To UBound(VarArray)
            VarArray(site) = DstVar
        Next site
        DstVar = VarArray
        ConvertVariableToArray = True
    End If

End Function

Private Sub GetChanList(ByVal PinList As String, ByVal site As Long, ByVal chanType As chtype, ByRef retChannels() As Long)

    Dim ChanNum As Long
    Dim siteNum As Long
    Dim errMsg As String

    Call TheExec.DataManager.GetChanList(PinList, site, chanType, retChannels, ChanNum, siteNum, errMsg)
    If errMsg <> "" Then
        Call DebugMsg(errMsg & " (at GetChanList)")
    End If
End Sub

Private Sub GetActiveChanList(ByVal PinList As String, ByVal chanType As chtype, ByRef retChannels() As Long)

    Dim ChanNum As Long
    Dim siteNum As Long
    Dim errMsg As String

    Call TheExec.DataManager.GetChanListForSelectedSites(PinList, chanType, retChannels, ChanNum, siteNum, errMsg)
    If errMsg <> "" Then
        Call DebugMsg(errMsg & " (at GetActiveChanList)")
    End If
End Sub

Private Function CountExistSite() As Long
    CountExistSite = TheExec.sites.ExistingCount - 1
End Function

Private Function CountActiveSite() As Long

    With TheExec.sites
        If .InSerialLoop Then
            CountActiveSite = 1
        Else
            CountActiveSite = .ActiveCount
        End If
    End With

End Function

Private Function IsActiveSite(ByVal site As Long) As Boolean
    IsActiveSite = TheExec.sites.site(site).Selected
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

Private Sub SeparatePinList(ByVal PinList As String, ByRef retPinNames() As String)

    Dim pinNum As Long
    Call TheExec.DataManager.DecomposePinList(PinList, retPinNames, pinNum)

End Sub

Private Function CreateEmpty2DArray(ByVal Dim1 As Long, ByVal Dim2 As Long) As Variant

    Dim ret2DArr() As Variant
    Dim tmp() As Double
    Dim i As Long

    ReDim ret2DArr(Dim1)
    ReDim tmp(Dim2)

    For i = 0 To UBound(ret2DArr)
        ret2DArr(i) = tmp
    Next i

    CreateEmpty2DArray = ret2DArr

End Function

Private Function IsValidSite(ByVal site As Long) As Boolean

    If site = ALL_SITE Then
        IsValidSite = True
    ElseIf 0 <= site And site <= CountExistSite Then
        IsValidSite = True
    Else
        IsValidSite = False
    End If
End Function
'################################################################################

'#################################### Debug #####################################









Public Sub DebugMsg(ByVal Msg As String)

    Msg = "Error Message: " & vbCrLf & "    " & Msg & vbCrLf
    Msg = Msg & "Test Instance Name: " & vbCrLf & "    " & TheExec.DataManager.InstanceName

    Call MsgBox(Msg, vbExclamation Or vbOKOnly, "Error")

End Sub
'################################################################################

'################################# Error Check ##################################
Private Function CheckPinList(ByVal PinList As String, ByVal chanType As chtype, ByVal FunctionName As String) As Boolean

    If GetChanType(PinList) <> chanType Then
        Call DebugMsg(PinList & " is Invalid Channel Type at " & FunctionName & "().")
        CheckPinList = False
    Else
        CheckPinList = True
    End If

End Function

Private Function CheckSinglePins(ByVal PinList As String, ByVal chanType As chtype, ByVal FunctionName As String) As Boolean

    Dim Channels() As Long
    Dim ChanNum As Long
    Dim siteNum As Long
    Dim errMsg As String
    Call TheExec.DataManager.GetChanList(PinList, ALL_SITE, chanType, Channels, ChanNum, siteNum, errMsg)

    If ChanNum <> siteNum Then
        Call DebugMsg("Don't Support Multi Pins at " & FunctionName & "().")
        CheckSinglePins = False
    Else
        CheckSinglePins = True
    End If

End Function

Private Function CheckForceVariantValue(ByVal ForceVal As Variant, ByVal loLim As Double, ByVal hiLim As Double, ByVal FunctionName As String) As Boolean

    Dim site As Long

    If IsArray(ForceVal) Then
        If UBound(ForceVal) <> CountExistSite Then
            Call DebugMsg("ForceVal is Invalid Site Array at " & FunctionName & "().")
            CheckForceVariantValue = False
            Exit Function
        End If

        For site = 0 To CountExistSite
            If (ForceVal(site) < loLim Or hiLim < ForceVal(site)) Then
                Call DebugMsg("ForceVal(= " & ForceVal(site) & ") must be between " & loLim & " and " & hiLim & " at" & FunctionName & "().")
                CheckForceVariantValue = False
                Exit Function
            End If
        Next site

    Else
        If (ForceVal < loLim Or hiLim < ForceVal) Then
            Call DebugMsg("ForceVal(= " & ForceVal & ") must be between " & loLim & " and " & hiLim & " at" & FunctionName & "().")
            CheckForceVariantValue = False
            Exit Function
        End If
    End If

    CheckForceVariantValue = True

End Function

Private Function CheckClampValue(ByVal clampVal As Double, ByVal loLim As Double, ByVal hiLim As Double, ByVal FunctionName As String) As Boolean

    If (clampVal < loLim Or hiLim < clampVal) Then
        Call DebugMsg("ClampVal(= " & clampVal & ") must be between " & loLim & " and " & hiLim & " at" & FunctionName & "().")
        CheckClampValue = False
    Else
        CheckClampValue = True
    End If

End Function

Private Function IsExistSite(ByVal site As Long, ByVal FunctionName As String) As Boolean

    If site <> ALL_SITE And (site < 0 Or CountExistSite < site) Then
        Call DebugMsg("Site(= " & site & ") must be -1 or between 0 and " & CountExistSite & " at" & FunctionName & "().")
        IsExistSite = False
    Else
        IsExistSite = True
    End If

End Function

Private Function CheckResultArray(ByRef retResult() As Double, ByVal FunctionName As String) As Boolean

    If UBound(retResult) <> CountExistSite Then
        Call DebugMsg("Elements of retResult() is Different from Number of Site at" & FunctionName & "().")
        CheckResultArray = False
    Else
        CheckResultArray = True
    End If

End Function

Private Function CheckAvgNum(ByVal avgNum As Long, ByVal FunctionName As String) As Boolean

    If avgNum < 1 Then
        Call DebugMsg("AvgNum must be 1 or More at" & FunctionName & "().")
        CheckAvgNum = False
    Else
        CheckAvgNum = True
    End If

End Function

Private Function CreateLimit(ByVal dstVal As Variant, ByVal loLim As Double, ByVal hiLim As Double) As Variant

    Dim i As Long

    If IsArray(dstVal) Then
        For i = 0 To UBound(dstVal)
            If dstVal(i) < loLim Then dstVal(i) = loLim
            If dstVal(i) > hiLim Then dstVal(i) = hiLim
        Next i
    Else
        If dstVal < loLim Then dstVal = loLim
        If dstVal > hiLim Then dstVal = hiLim
    End If

    CreateLimit = dstVal

End Function
'################################################################################

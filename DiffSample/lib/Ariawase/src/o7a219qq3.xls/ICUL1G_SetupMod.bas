Attribute VB_Name = "ICUL1G_SetupMod"
Option Explicit
'  2013/10/23 H.Arikawa GetCVS_v4_0を入れ込み。

Public Type MIPI_SETUP
    MipiKeyName As String
    Threshold_Board16 As Double
    Threshold_Board19 As Double
    UserDelayCLK(nSite) As Long
    UserDelay00(nSite) As Long
    UserDelay01(nSite) As Long
    UserDelay02(nSite) As Long
    UserDelay03(nSite) As Long
    VodSetCLK(nSite) As Double
    VodSet00(nSite) As Double
    VodSet01(nSite) As Double
    VodSet02(nSite) As Double
    VodSet03(nSite) As Double
End Type

'++++++++++ DIGITAL OUT SELECT MODE ++++++++++ 2013/02/04
Public Enum EnumDOUT
    CMOS = 0
    LVDS_1CH = 1
    LVDS_2CH = 2
    LVDS_2CH_CRC = 3
    LVDS_2CH_DATA_CLOCK = 4
    LVDS_2CH_DATA_CLOCK_CRC = 5
    LVDS_1CH_COMP8_CRC = 6
    LVDS_1CH_RAW10_CRC = 7
    MIPI_2CH_RAW10_CRC = 8
    MIPI_1CH_RAW10_CRC = 9
    MIPI_2CH_RAW10_DC = 10
    MIPI_4CH_RAW10_CRC = 11
    MIPI_3CH_RAW10_CRC = 12
End Enum

Public CapFreq_Typ As Double
Public MipiSetFor1G(20) As MIPI_SETUP

Private Const ARG_MIPI_CAPTURE_PIN_NAME As Long = 10              'テストインスタンスのArg30

Public MIPI_DCKP_NAME As String
Public MIPI_DO0P_NAME As String
Public MIPI_DO1P_NAME As String
Public MIPI_DO2P_NAME As String
Public MIPI_DO3P_NAME As String

'MipiKeyNameのチェック
Public Function ICUL1G_Parameter_Check(ByVal MipiKeyName As String) As Boolean
    Dim intParamNo  As Integer
    
    If MipiKeyName = "" Then
        MsgBox "Search Error! [" & MipiKeyName & "] is Empty!"
        Exit Function
    End If
    
    For intParamNo = 0 To UBound(MipiSetFor1G)
        '既に登録されていたら抜ける
        If MipiSetFor1G(intParamNo).MipiKeyName = MipiKeyName Then
            ICUL1G_Parameter_Check = True
            Exit For
        ElseIf MipiSetFor1G(intParamNo).MipiKeyName = "" Then
        '空き領域に達したら、MipiKeyNameを設定
            MipiSetFor1G(intParamNo).MipiKeyName = MipiKeyName
            ICUL1G_Parameter_Check = False
            Exit For
        End If
    Next intParamNo

End Function

Public Sub ICUL1G_Parameter_Def()
    Const ShtAcqTbl   As String = "TestCondition"
    Const strAcquireKey As String = "FW_SetICUL1G"
    Const intArgNumKey As Integer = 2
    Dim MipiKeyName As String

    '変数クリア
    Erase MipiSetFor1G
        
    Dim wkshtObj As Object
    
    '======= WorkSheet Select ========
    Set wkshtObj = ThisWorkbook.Sheets(ShtAcqTbl)
         
    '======= WorkSheet ErrorProcess ========
    If IsEmpty(wkshtObj) Then
        MsgBox "Non [" & ShtAcqTbl & "] WorkSheet"
        Exit Sub
    End If

    Dim StartPoint_Row As Long
    Dim StartPoint_Column As Long
    Dim intSearchRow As Long
    
    StartPoint_Row = 5                                                  '2013/02/07 H.Arikawa 修正
    StartPoint_Column = 3
    
    For intSearchRow = StartPoint_Row To 32000
        If wkshtObj.Cells(intSearchRow, StartPoint_Column) = "" Then
            Exit For
        ElseIf wkshtObj.Cells(intSearchRow, StartPoint_Column) = strAcquireKey Then
            MipiKeyName = wkshtObj.Cells(intSearchRow, StartPoint_Column + intArgNumKey + 1)    '2013/02/07 H.Arikawa 修正
            If ICUL1G_Parameter_Check(MipiKeyName) = False Then
                Call GetCSVData(sub_CSV_CTRL(strPassForCSV & MipiKeyName & "_" & Format(CStr(Sw_Node), "000") & ".csv"))
                Call CopyCSVData(MipiKeyName)
                '---- UserDelay Setting --------
                Call get_idelay(MipiKeyName)
                '---- Threshold Setting --------
                Call get_Threshold(MipiKeyName)
                '---- VOD Setting --------
                Call get_VOD(MipiKeyName)
            End If
        End If
    Next intSearchRow

End Sub

'MipiKeyNameシートから、ReadCSVシートにコピー
Public Sub CopyCSVData(ByVal MipiKeyName As String)
    
    Dim copyRow As Integer
    Dim copyCol As Integer
    
    For copyRow = 1 To 50
        For copyCol = 1 To 50
            Worksheets(MipiKeyName).Cells(copyRow, copyCol).Value = Worksheets("Read CSV").Cells(copyRow, copyCol).Value
        Next
    Next

End Sub

'MipiKeyNameから、設定されている番号を返す
Public Function getMipiNum(ByVal MipiKeyName As String) As Integer

    Dim intParamNo  As Integer
    
    For intParamNo = 0 To UBound(MipiSetFor1G)
        If MipiSetFor1G(intParamNo).MipiKeyName = MipiKeyName Then
            getMipiNum = intParamNo
            Exit Function
        End If
    Next intParamNo

    MsgBox "Search Error! Not Finding MIPI Parameter[" & MipiKeyName & "]"

End Function

'2013/02/12 H.Arikawa 処理修正
Public Function set_VODLevel(ByVal strOutLinePins As String, ByVal MipiKeyName As String) As Long

    Dim DCKChans() As Long
    Dim DO3Chans() As Long
    Dim DO2Chans() As Long
    Dim DO1Chans() As Long
    Dim DO0Chans() As Long

    Dim nChans As Long
    Dim nSites As Long
    Dim errMsg As String
    Dim site As Long

    With TheExec.DataManager
        Call .GetChanList(MIPI_DCKP_NAME, -1, chICUL1Gclk_high, DCKChans(), nChans, nSites, errMsg)
        Call .GetChanList(MIPI_DO0P_NAME, -1, chICUL1Gdata_high, DO0Chans(), nChans, nSites, errMsg)
        If (strOutLinePins = "MIPI_2LANE") Or (strOutLinePins = "MIPI_4LANE") Then
            Call .GetChanList(MIPI_DO1P_NAME, -1, chICUL1Gdata_high, DO1Chans(), nChans, nSites, errMsg)
        End If
        If (strOutLinePins = "MIPI_4LANE") Then
            Call .GetChanList(MIPI_DO2P_NAME, -1, chICUL1Gdata_high, DO2Chans(), nChans, nSites, errMsg)
            Call .GetChanList(MIPI_DO3P_NAME, -1, chICUL1Gdata_high, DO3Chans(), nChans, nSites, errMsg)
        End If
    End With

    With MipiSetFor1G(getMipiNum(MipiKeyName))
        For site = 0 To nSite
            TheHdw.ICUL1G.Chans(DCKChans(site), icul1gChannelTypeClkHigh).PinLevels.ModifyDiffLevel icul1gPinLevelVod, .VodSetCLK(site), icul1gVtOpModeDynamic
            TheHdw.ICUL1G.Chans(DO0Chans(site), icul1gChannelTypeDataHigh).PinLevels.ModifyDiffLevel icul1gPinLevelVod, .VodSet00(site), icul1gVtOpModeDynamic
            TheHdw.ICUL1G.Chans(DCKChans(site), icul1gChannelTypeClkHigh).PinLevels.ModifyDiffLevel icul1gPinLevelVt, 0.2 * V, icul1gVtOpModeDynamic                    'ICUL1G Vod設定不具合対策(PATCHで改善だが、念の為)
            TheHdw.ICUL1G.Chans(DO0Chans(site), icul1gChannelTypeDataHigh).PinLevels.ModifyDiffLevel icul1gPinLevelVt, 0.2 * V, icul1gVtOpModeDynamic                   'ICUL1G Vod設定不具合対策(PATCHで改善だが、念の為)
            If (strOutLinePins = "MIPI_2LANE") Or (strOutLinePins = "MIPI_4LANE") Then
                TheHdw.ICUL1G.Chans(DO1Chans(site), icul1gChannelTypeDataHigh).PinLevels.ModifyDiffLevel icul1gPinLevelVod, .VodSet01(site), icul1gVtOpModeDynamic
                TheHdw.ICUL1G.Chans(DO1Chans(site), icul1gChannelTypeDataHigh).PinLevels.ModifyDiffLevel icul1gPinLevelVt, 0.2 * V, icul1gVtOpModeDynamic               'ICUL1G Vod設定不具合対策(PATCHで改善だが、念の為)
            End If
            If (strOutLinePins = "MIPI_4LANE") Then
                TheHdw.ICUL1G.Chans(DO2Chans(site), icul1gChannelTypeDataHigh).PinLevels.ModifyDiffLevel icul1gPinLevelVod, .VodSet02(site), icul1gVtOpModeDynamic
                TheHdw.ICUL1G.Chans(DO3Chans(site), icul1gChannelTypeDataHigh).PinLevels.ModifyDiffLevel icul1gPinLevelVod, .VodSet03(site), icul1gVtOpModeDynamic
                TheHdw.ICUL1G.Chans(DO2Chans(site), icul1gChannelTypeDataHigh).PinLevels.ModifyDiffLevel icul1gPinLevelVt, 0.2 * V, icul1gVtOpModeDynamic               'ICUL1G Vod設定不具合対策(PATCHで改善だが、念の為)
                TheHdw.ICUL1G.Chans(DO3Chans(site), icul1gChannelTypeDataHigh).PinLevels.ModifyDiffLevel icul1gPinLevelVt, 0.2 * V, icul1gVtOpModeDynamic               'ICUL1G Vod設定不具合対策(PATCHで改善だが、念の為)
            End If
        Next site
    End With

End Function


'===============================================================================================
'Test Condition Sheet
'===============================================================================================
Public Sub FW_SetICUL1G(ByVal Parameter As CSetFunctionInfo)
    
    Dim site As Long '2013/02/04

    If Flg_Simulator = 1 Then Exit Sub

    '++++ PMD SELECT ++++++++++++++++++++++++++++++++++++++
    Dim strOutLinePins As String
    Dim lngACQ_Width As Long
    Dim MipiKeyName As String

    With Parameter
        strOutLinePins = .Arg(0)
        lngACQ_Width = .Arg(1)
        MipiKeyName = .Arg(2)
    End With

    '++++ ICUL1G SetUp ++++++++++++++++++++++++++++++++++++
    TheHdw.ICUL1G.Pins("MIPI_4LANE").Disconnect 'MAX-Lane Disconnect

    '++++ ICUL1G Connect +++++++++++++++++++++++++++++++++++++++++
    TheHdw.ICUL1G.Pins(strOutLinePins).Connect                        'ICUL1G 不具合対策(PATCHで改善だが、念の為)

    '++++++++++++     Set Vod        +++++++++++
    Call set_VODLevel(strOutLinePins, MipiKeyName)
    '++++++++++++     Set Threshold  +++++++++++

    With TheHdw.ICUL1G
        Call .board(16).ModifyMipiLpThreshold(MipiSetFor1G(getMipiNum(MipiKeyName)).Threshold_Board16)
        Call .board(19).ModifyMipiLpThreshold(MipiSetFor1G(getMipiNum(MipiKeyName)).Threshold_Board19)

        Select Case UCase(Parameter.Arg(3))
            Case "FULLSTATE", ""
                .MipiCaptureMode = icul1gMipiLpFullStateSync
            Case "11STATE"
                .MipiCaptureMode = icul1gMipiLp11StateSync
            Case "IGNORE"
                .MipiCaptureMode = icul1gMipiLpStateIgnore
            Case Else
                MsgBox "Input Error! @Image ACQTBL Sheet! " & MipiKeyName & " : Prease Insert [FullState] or [11State] or [Ignore] in the [Arg8] " '2013/02/04
        End Select
    End With
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            With TheHdw.ICUL1G.ImgPins(strOutLinePins, site)
                Select Case UCase(Parameter.Arg(4))
                    Case "RAW10", ""
                        .MIPI.DataFormat = icul1gDataFormatMipiCsi2Raw10
                    Case "RAW12"
                        .MIPI.DataFormat = icul1gDataFormatMipiCsi2Raw12
                    Case "RAW8"
                        .MIPI.DataFormat = icul1gDataFormatMipiCsi2Raw8
                    Case "RGB444"
                        .MIPI.DataFormat = icul1gDataFormatMipiCsi2RGB444
                    Case "RGB555"
                        .MIPI.DataFormat = icul1gDataFormatMipiCsi2RGB555
                    Case "RGB565"
                        .MIPI.DataFormat = icul1gDataFormatMipiCsi2RGB565
                    Case "YUV422Bit10"
                        .MIPI.DataFormat = icul1gDataFormatMipiCsi2YUV422Bit10
                    Case "YUV422Bit8"
                        .MIPI.DataFormat = icul1gDataFormatMipiCsi2YUV422Bit8
                    Case Else
                        MsgBox "Input Error! @Image ACQTBL Sheet! " & MipiKeyName & " : Prease Insert [Row8] or [Row10] or [Row12] or [RGB444] or [RGB555] or [RGB565] or [YUV422Bit10] or [YUV422Bit8] in the [Arg4] "  '2013/02/04
                End Select
                Select Case UCase(Parameter.Arg(5))
                    Case "ENABLE"
                        .MIPI.SyncCodeCapture = icul1gSyncCodeCaptureEnabled
                    Case "DISABLE"
                        .MIPI.SyncCodeCapture = iculSyncCodeCaptureDisabled
                    Case Else
                        MsgBox "Input Error! @TestCondition! " & MipiKeyName & " : Prease Insert [Enable] or [Disable] in the [Arg5] "  '2013/02/04
                End Select

                .FrameSkip = CLng(Parameter.Arg(6))
                .LineSkip = CLng(Parameter.Arg(7))

                .UserDelay(icul1gClockStrobeLane) = MipiSetFor1G(getMipiNum(MipiKeyName)).UserDelayCLK(site)
                .UserDelay(icul1gDataLane00) = MipiSetFor1G(getMipiNum(MipiKeyName)).UserDelay00(site)
                If strOutLinePins = "MIPI_2LANE" Or strOutLinePins = "MIPI_4LANE" Then
                    .UserDelay(icul1gDataLane01) = MipiSetFor1G(getMipiNum(MipiKeyName)).UserDelay01(site)
                End If
                If strOutLinePins = "MIPI_4LANE" Then
                    .UserDelay(icul1gDataLane02) = MipiSetFor1G(getMipiNum(MipiKeyName)).UserDelay02(site)
                    .UserDelay(icul1gDataLane03) = MipiSetFor1G(getMipiNum(MipiKeyName)).UserDelay03(site)
                End If
                .ImageDataWidth = lngACQ_Width '同期コードを取り込む設定でも同期コード無しの幅を設定する。
                .ClearAlarm (icul1gAlarmCaptureRelated)
                .alarm(icul1gAlarmCaptureRelated) = icuAlarmOff
            End With
        End If
    Next site

    'MIPI RreqJudge_Typ[MHz]
    CapFreq_Typ = Parameter.Arg(8)

End Sub

'InitTest内でコールされる
Public Sub InitializeCaptureUnitInside()

    If TheExec.TesterMode = testModeOffline Then Exit Sub
    
    '===== CAPTURE UNIT INITIALIZE ========================
    Application.Run ("ICUL1G_Parameter_Def")            'CaptureシステムがICUL1Gの時に実行
    Application.Run ("MipiCapturePinNameGet")           'CapturePin情報を取得
    
End Sub

'InitTest外でコールされる
Public Sub InitializeCaptureUnitOutSide()

    If TheExec.TesterMode = testModeOffline Then Exit Sub
    
    '===== CAPTURE UNIT INITIALIZE ========================
    Application.Run ("ICUL1G_InitializeCaptureUnitOutSide")            'CaptureシステムがICUL1Gの時に実行
    
End Sub
Public Sub ICUL1G_InitializeCaptureUnitOutSide()
    '====POWER ON========
    TheHdw.PinLevels.ApplyPower                             'Capture MUST
    TheHdw.IDP.CaptureTimeOut = 1                           'Capture MUST
End Sub

Public Sub CaptureResetSequence()  '2013/02/04

    If TheExec.TesterMode = testModeOffline Then Exit Sub
    
    '===== CAPTURE UNIT INITIALIZE ========================
    Application.Run ("ICUL1G_CaptureResetSequence")            'CaptureシステムがICUL1Gの時に実行

End Sub

Public Sub ICUL1G_CaptureResetSequence()  '2013/02/04

    TheHdw.IDP.WaitCaptureCompletion
    TheHdw.IDP.WaitTransferCompletion
    TheHdw.IDP.ResetMultiCapture

End Sub

Public Sub MipiCapturePinNameGet()  '2013/02/28 TestInstancesの出力仕様に合わせて更新
    
    'Test Instanceシートからパラメータを取得。
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, -1) Then
        Err.Raise 9999, "PutImageInto_Common", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
    End If

    'ARG_MIPI_CAPTURE_PIN_NAMEから開始して、","でSplitして"NULL"になるところまでPin取得する処理を行う。

    Dim si As Integer

    Dim tmpCapturePinName(10) As String
    Dim tmpRevCapturePinName() As String
    Dim tmpName As String
    Dim tmpNameCheck As String
    Dim tmpCapturePinNameCheck As String
    Dim k As Integer
    Dim j As Integer
    Dim checknum As Integer
    
    j = 0
    
    checknum = UBound(ArgArr)
    
    For si = ARG_MIPI_CAPTURE_PIN_NAME To checknum
        If ArgArr(si) = "," Or ArgArr(si) = "" Then
            GoTo NextCheck
        Else:
        tmpName = ArgArr(si)
        tmpRevCapturePinName = Split(tmpName, ",")
            For k = 0 To UBound(tmpRevCapturePinName)
                tmpCapturePinNameCheck = Left(tmpRevCapturePinName(k), 3)
                If tmpCapturePinNameCheck = "Ph_" Then
                    GoTo NextPhCheck
                Else:
                    tmpCapturePinName(j) = tmpRevCapturePinName(k)
                    Select Case j
                        Case 0
                            MIPI_DCKP_NAME = tmpCapturePinName(j)
                        Case 1
                            MIPI_DO0P_NAME = tmpCapturePinName(j)
                        Case 2
                            MIPI_DO1P_NAME = tmpCapturePinName(j)
                        Case 3
                            MIPI_DO2P_NAME = tmpCapturePinName(j)
                        Case 4
                            MIPI_DO3P_NAME = tmpCapturePinName(j)
                        Case Else
                    End Select
                    j = j + 1
                End If
NextPhCheck:
            Next k
        End If
NextCheck:
    Next si

End Sub

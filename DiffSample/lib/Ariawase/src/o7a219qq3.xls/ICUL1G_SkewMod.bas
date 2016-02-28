Attribute VB_Name = "ICUL1G_SkewMod"
Option Explicit
'=== 自動化未デバッグ 現行JOBからCopyしただけ コンパイルは問題無し　2013/09/20 H.Arikawa ===

'==========================================================================
'   ICUL1G_SkewMod.bas
'   Fuction : This module is skew map tool for ICUL1G capture pins.
'
'   First Created By TAG(Teradyne)  2011/07/28
'
'   ///Version 2.1///
'
'   Update history
'   Draft1.0    first edition for debugging.
'   Ver1.0      modify Skew tool, and adding adjustment function of LPS-threshold. by 2011/08/04
'   Ver1.1      Correction of bug (Vod setting routine and Min/Max comparison of skew of Skew tool). by 2011/08/05
'   Ver1.2      Correction of bug (changed blAlarmCaptureRelated type form long to boolean, etc). by 2011/08/11
'   Ver2.0      modify Skew and LPS-threshold tool (adding comparison area of expect image is specified, etc). by 2011/08/19
'   Ver2.1      modify Skew tool (FreqCnt is deleted from judge and output format change to seat). by 2011/09/12
'==========================================================================


Public Function ICUL1G_Capture_Skew(ByRef sICUL1GPins As String, ByRef sCapPlane As CImgPlane, sCapPad As String, lStartTap As Long, lStopTap As Long, lStepTap As Long, _
                                    Optional CheckTP As Boolean = False, Optional sExptPlane As String = "", Optional sWorkPlane As String = "", Optional sNonCompPad As String = "", Optional sExptFile As String = "", _
                                    Optional Vod_flg As Boolean = False, Optional dMaxVod As Double = 0, Optional dMinVod As Double = 0, Optional dStepVod As Double = 1) As Long
    
    Dim cPins(0) As String
    Dim cDest(0) As String
    Dim lPortNum As Long
    Dim lngRetMatchVal(12) As Long
    Dim lngUnMatchCnt() As Long
    Dim my_Vod() As Double
    Dim my_Tap() As Long
    Dim my_MidTap() As Long
    Dim my_MinTap() As Long
    Dim my_MaxTap() As Long
    Dim my_MarFlg() As Long
    Dim lWidthTap As Long
    Dim my_RunModeFlg As Boolean
    Dim my_CapTimeOut As Double
    Dim strVodInfo As String
    Dim strTapInfo As String
    Dim lSiteNum As Long
    Dim lMaxSite As Long
    Dim lNumLane As Long
    Dim lngPinStdd As Long
    
    'Frequency Counter
    Dim lngFreqCnt() As Long
    Dim dblFreq() As Long
    'Capture Result
    Dim capFlag() As Double
    Dim lngCapturedFrames() As Long
    Dim lngReadAcqFlag() As Long
    'Alarm
    Dim blAlarmCaptureRelated() As Boolean      'Changes with Ver1.2
    
    Dim lngClockEstDelay() As Long
    Dim lngDataEstDelay() As Long
    Dim SheetName As String
    Dim skew_channel As Long
    Dim next_ch_cells As Long
    Dim vod_max As Double
    Dim vod_min As Double
    Dim vod_resolution As Double
    Dim vod_val As Double
    Dim vod_max_step As Long
    Dim vod_cnt As Long
    Dim rows_step As Long
    Dim column_step As Long
    Dim site_ofst As Long                       'Adds with Ver2.1
    Dim site_cnt As Long                        'Adds with Ver2.1
    Dim jdg_color As Long                       'Adds with Ver2.1
    
    Dim nMin As Long
    Dim nMax As Long
    Dim dataLaneNum As Long
    Dim lngUserDelay As Long
    Dim soutaiflag As Long
    Dim trimflag As Long
    
    Dim MeasChans() As Long
    Dim nChans As Long
    Dim nSites As Long
    Dim errMsg As String
    Dim SetChans() As Long
    Dim ChArryFlg As Long
    Dim lngChArryNum As Long
    Dim strMipiCaptureMode As String            'Adds with Ver2.0
    Dim lng1gChType As Icul1gChannelType

    lMaxSite = TheExec.sites.ExistingCount
    lNumLane = TheHdw.ICUL1G.ImgPins(sICUL1GPins).DataLineNum
    ReDim my_Vod(lMaxSite - 1, lNumLane) As Double
    ReDim my_Tap(lMaxSite - 1) As Long
    ReDim my_MidTap(lMaxSite - 1, lNumLane) As Long
    ReDim my_MinTap(lMaxSite - 1) As Long
    ReDim my_MaxTap(lMaxSite - 1) As Long
    ReDim SetChans(lMaxSite - 1) As Long
    'Frequency Counter
    ReDim lngFreqCnt(lMaxSite - 1) As Long
    ReDim dblFreq(lMaxSite - 1) As Long
    'Capture Result
    ReDim capFlag(lMaxSite - 1)
    ReDim lngCapturedFrames(lMaxSite - 1)
    ReDim lngReadAcqFlag(lMaxSite - 1)
    ReDim lngUnMatchCnt(lMaxSite - 1)
    'Alarm
    ReDim blAlarmCaptureRelated(lMaxSite - 1)   'Changes with Ver1.2
    
    soutaiflag = 0  'check
    trimflag = 0    'check
    

    my_RunModeFlg = False
    If TheExec.RunMode = runModeDebug Then
        TheExec.RunMode = runModeProduction
        my_RunModeFlg = True
    End If
    my_CapTimeOut = TheHdw.IDP.CaptureTimeOut
    TheHdw.IDP.CaptureTimeOut = 0.3

    If Vod_flg = False Then
        vod_resolution = 1
        vod_max = 0
        vod_min = 0
        next_ch_cells = 6
        SheetName = "Skew"
    ElseIf Vod_flg = True Then
        vod_resolution = dStepVod
        vod_max = dMaxVod
        vod_min = dMinVod
        next_ch_cells = 22
        SheetName = "SkewVod"
    End If
    vod_cnt = 0
    vod_max_step = ((vod_max - vod_min) / (vod_resolution)) + 2
    
    TheHdw.Raw.icul1g_drv.ICUL1G_GetStandard sICUL1GPins, lngPinStdd
    If lngPinStdd = 1 Then                      'Adds with Ver2.0
        If TheHdw.ICUL1G.MipiCaptureMode = 0 Then
            strMipiCaptureMode = "icul1gMipiLpFullStateSync"
        ElseIf TheHdw.ICUL1G.MipiCaptureMode = 1 Then
            strMipiCaptureMode = "icul1gMipiLp11StateSync"
        Else
            strMipiCaptureMode = "icul1gMipiLpStateIgnore"
        End If
    Else
        strMipiCaptureMode = ""
    End If
    
    If Vod_flg = True Then                      'Changes with Ver2.1
        For lSiteNum = 0 To lMaxSite - 1
            If TheExec.sites.site(lSiteNum).Active = True Then
                Worksheets(SheetName & lSiteNum).Cells.Clear
                Worksheets(SheetName & lSiteNum).Cells.Interior.ColorIndex = xlNone
            End If
         Next lSiteNum
    Else
        Worksheets(SheetName).Cells.Clear
        Worksheets(SheetName).Cells.Interior.ColorIndex = xlNone
    End If
    
    'Initialize Alarm
    With TheHdw.ICUL1G.ImgPins(sICUL1GPins)
        .alarm(icul1gAlarmAll) = icuAlarmRedirect
        .ClearAlarm (icul1gAlarmAll)
    End With
    
    'Capture setup
    cPins(0) = sICUL1GPins
    cDest(0) = sCapPlane.Name
    lPortNum = 1
'    TheHdw.IDP.SetPMD cDest(0), sCapPad
    TheHdw.IDP.WritePixel cDest(0), idpColorFlat, 0
    
    'Load Reference Image
    If CheckTP = True Then
        TheHdw.IDP.SetPMD sExptPlane, sCapPad
        TheHdw.IDP.WritePixel sExptPlane, idpColorFlat, 0
        For lSiteNum = 0 To lMaxSite - 1
            If TheExec.sites.site(lSiteNum).Active Then
                TheExec.Datalog.WriteComment "Loading expect image of site" & lSiteNum & "...  " & sExptFile & " to " & sExptPlane
                TheHdw.IDP.ReadFile lSiteNum, sExptPlane, idpColorFlat, sExptFile, idpFileBinary
            End If
        Next lSiteNum
    End If

    '----------------------------------------------------------------------------------------
    skew_channel = 0
    For dataLaneNum = -1 To lNumLane - 1
        Call TheExec.DataManager.GetChanList(sICUL1GPins, -1, chAll, MeasChans(), nChans, nSites, errMsg)
        ChArryFlg = 0
        For lSiteNum = 0 To lMaxSite - 1
            If TheExec.sites.site(lSiteNum).Active = True Then
                If dataLaneNum = -1 Then
                    lngChArryNum = lSiteNum
                Else
                    lngChArryNum = (lNumLane - dataLaneNum) * lMaxSite * 2 + lSiteNum
                End If
                If ((MeasChans(lngChArryNum) >= 0) And (MeasChans(lngChArryNum) <= 7)) Or _
                   ((MeasChans(lngChArryNum) >= 24) And (MeasChans(lngChArryNum) <= 31)) Or _
                   ((MeasChans(lngChArryNum) >= 48) And (MeasChans(lngChArryNum) <= 55)) Or _
                   ((MeasChans(lngChArryNum) >= 72) And (MeasChans(lngChArryNum) <= 79)) Then
                    lng1gChType = icul1gChannelTypeClkHigh
                Else
                    lng1gChType = icul1gChannelTypeDataHigh
                End If
                my_Vod(lSiteNum, dataLaneNum + 1) = TheHdw.ICUL1G.Chans(MeasChans(lngChArryNum), lng1gChType).PinLevels.ReadDiffPinLevels(icul1gPinLevelVod)
                SetChans(ChArryFlg) = MeasChans(lngChArryNum)
                ChArryFlg = ChArryFlg + 1
            End If
        Next lSiteNum
        
        For vod_val = vod_max To vod_min Step vod_resolution
            If Vod_flg = True Then
                Call TheHdw.ICUL1G.Chans(SetChans(), lng1gChType).PinLevels.ModifyDiffLevel(icul1gPinLevelVod, vod_val)
                vod_cnt = vod_cnt + 1
            End If
            rows_step = (dataLaneNum + 3) + (next_ch_cells * skew_channel + vod_max_step * skew_channel) + vod_cnt
        
            Call TheHdw.ICUL1G.ImgPins(sICUL1GPins).GetUserDelayRange(dataLaneNum, nMin, nMax)
            If lngPinStdd = 2 Then
                TheHdw.ICUL1G.ImgPins(sICUL1GPins).serial.EstimateEdgeAlignment 400# * MHz, icul1gClockTypeSerialDDR, lngClockEstDelay(), lngDataEstDelay()
            Else
                ReDim lngClockEstDelay(lMaxSite - 1)
                ReDim lngDataEstDelay(lMaxSite - 1)
                For lSiteNum = 0 To lMaxSite - 1
                    lngClockEstDelay(lSiteNum) = 0
                    lngDataEstDelay(lSiteNum) = 0
                Next lSiteNum
            End If
                        
            site_cnt = 0                                    'Adds with Ver2.1
            For lSiteNum = 0 To lMaxSite - 1
                site_ofst = (lNumLane + 4) * site_cnt       'Adds with Ver2.1
                If trimflag = 1 Then site_ofst = site_ofst * 2
                If TheExec.sites.site(lSiteNum).Active = True Then
                    my_Tap(lSiteNum) = TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum)
                    my_MinTap(lSiteNum) = lStopTap
                    my_MaxTap(lSiteNum) = lStartTap
                    If dataLaneNum = -1 Then
                        If Vod_flg = True Then
                            strVodInfo = "Clock: Vod=" & Round(vod_val * 1000, 2) & "[mV]"
                        Else
                            strVodInfo = "Clock: Vod=" & Round(my_Vod(lSiteNum, dataLaneNum + 1) * 1000, 2) & "[mV]"
                        End If
                        strTapInfo = "Tap Setting Value = " & my_Tap(lSiteNum) - lngClockEstDelay(lSiteNum)
                    Else
                        If Vod_flg = True Then
                            strVodInfo = "Lane" & dataLaneNum & ": Vod=" & Round(vod_val * 1000, 2) & "[mV]"
                        Else
                            strVodInfo = "Lane" & dataLaneNum & ": Vod=" & Round(my_Vod(lSiteNum, dataLaneNum + 1) * 1000, 2) & "[mV]"
                        End If
                        strTapInfo = "Tap Setting Value = " & my_Tap(lSiteNum) - lngDataEstDelay(lSiteNum)
                    End If
                    If Vod_flg = True Then              'Changes with Ver2.1
                        Worksheets(SheetName & lSiteNum).Cells(rows_step, 1).Value = strVodInfo
                        Worksheets(SheetName & lSiteNum).Cells(rows_step, 2).Value = strTapInfo
                    Else
                        Worksheets(SheetName).Cells(rows_step + site_ofst, 1).Value = strVodInfo
                        Worksheets(SheetName).Cells(rows_step + site_ofst, 2).Value = strTapInfo
                        site_cnt = site_cnt + 1
                    End If
                End If
            Next lSiteNum
            
            ReDim my_MarFlg(lMaxSite - 1) As Long
            For lngUserDelay = lStartTap To lStopTap Step lStepTap
                DoEvents
              
                site_cnt = 0                                    'Adds with Ver2.1
                For lSiteNum = 0 To lMaxSite - 1
                    site_ofst = (lNumLane + 4) * site_cnt       'Adds with Ver2.1
                    If trimflag = 1 Then site_ofst = site_ofst * 2
                    If TheExec.sites.site(lSiteNum).Active = True Then
                        If soutaiflag = 1 Then
                            TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum) = TheHdw.ICUL1G.ImgPins(sICUL1GPins).UserDelay(dataLaneNum) + lngUserDelay
                        Else
                            If dataLaneNum = -1 Then
                                If lngUserDelay <= nMax Then
                                    If lngUserDelay >= nMin Then
                                        TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum) = lngClockEstDelay(lSiteNum) + lngUserDelay
                                    Else
                                        TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum) = lngClockEstDelay(lSiteNum) + nMin
                                    End If
                                Else
                                    TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum) = lngClockEstDelay(lSiteNum) + nMax
                                End If
                            Else
                                If lngUserDelay <= nMax Then
                                    If lngUserDelay >= nMin Then
                                        TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum) = lngDataEstDelay(lSiteNum) + lngUserDelay
                                    Else
                                        TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum) = lngDataEstDelay(lSiteNum) + nMin
                                    End If
                                Else
                                    TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum) = lngDataEstDelay(lSiteNum) + nMax
                                End If
                            End If
                        End If
                        
                        If Vod_flg = True Then      'Changes with Ver2.1
                            Worksheets(SheetName & lSiteNum).Cells(1, 1).Value = "Site" & lSiteNum
                            Worksheets(SheetName & lSiteNum).Cells(1, 2).Value = strMipiCaptureMode     'Adds with Ver2.0
                            Worksheets(SheetName & lSiteNum).Cells(1, (lngUserDelay - lStartTap) / lStepTap + 3).Value = lngUserDelay
                        Else
                            Worksheets(SheetName).Cells(1 + site_ofst, 1).Value = "Site" & lSiteNum
                            Worksheets(SheetName).Cells(1 + site_ofst, 2).Value = strMipiCaptureMode    'Adds with Ver2.0
                            Worksheets(SheetName).Cells(1 + site_ofst, (lngUserDelay - lStartTap) / lStepTap + 3).Value = lngUserDelay
                            site_cnt = site_cnt + 1
                        End If
                    End If
                Next lSiteNum
            
                'ADD koga
                cPins(0) = sICUL1GPins
                cDest(0) = sCapPlane.Name
            
                'Start Capturing
                TheHdw.IDP.MultiAcquire cPins, cDest, lPortNum, 1, idpNonAverage, idpAcqNonInterlace, , , idpCurrentPmd
            
                'Wait for the image capture to complete
                TheHdw.IDP.WaitCaptureCompletion
                TheHdw.IDP.WaitTransferCompletion
    
                'Get CaputureResult
                For lSiteNum = 0 To lMaxSite - 1
                    If TheExec.sites.site(lSiteNum).Active = True Then
                        TheHdw.IDP.ReadAcquiredFrameCount lSiteNum, cDest(skew_channel), lngCapturedFrames(lSiteNum)
                        'Input AcquireStatus to FlagAry
                        TheHdw.IDP.ReadAcquireStatus lSiteNum, cDest(skew_channel), lngReadAcqFlag(lSiteNum)
                        'Get Capture Alarm Status
                        blAlarmCaptureRelated(lSiteNum) = TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).AlarmStatus(icul1gAlarmCaptureRelated)
                    End If
                Next lSiteNum
                
                If (CheckTP = True) And (sNonCompPad <> "") Then           'Adds with Ver2.0
                    'Written -1 in no comparison areas
                    With TheHdw.IDP
                        .SetPMD sCapPlane, sNonCompPad
                        .SetPMD sExptPlane, sNonCompPad
                        .WritePixel sCapPlane, idpColorFlat, -1
                        .WritePixel sExptPlane, idpColorFlat, -1
                    End With
                End If
                
                site_cnt = 0                                    'Adds with Ver2.1
                For lSiteNum = 0 To lMaxSite - 1
                    site_ofst = (lNumLane + 4) * site_cnt       'Adds with Ver2.1
                    If trimflag = 1 Then site_ofst = site_ofst * 2
                    If TheExec.sites.site(lSiteNum).Active = True Then
                        If CheckTP = True Then
                            'Matching and Jadge
                            With TheHdw.IDP
                                .SetPMD sCapPlane, sCapPad
                                .SetPMD sExptPlane, sCapPad
                                .SetPMD sWorkPlane, sCapPad
                                .WritePixel sWorkPlane, idpColorFlat, 0
                            
                                .LXor sCapPlane, idpColorFlat, sExptPlane, idpColorFlat, sWorkPlane, idpColorFlat
                                .Count sWorkPlane, idpColorFlat, idpCountOutside, 0, 0, idpLimitExclude
                                .ReadResult lSiteNum, sWorkPlane, idpColorFlat, idpReadCount, lngRetMatchVal
                            End With
                            lngUnMatchCnt(lSiteNum) = lngRetMatchVal(idpColorRed)
                        Else
                            lngUnMatchCnt(lSiteNum) = 0
                        End If
                        
                        If (lngReadAcqFlag(lSiteNum) <> idpAcqCompleted) Or (blAlarmCaptureRelated(lSiteNum) = True) Then   'Changes with Ver2.1
                            capFlag(lSiteNum) = 1
                        Else
                            capFlag(lSiteNum) = 0
                        End If
                        
                        If capFlag(lSiteNum) = 0 Then           'Changes with Ver2.1
                            If lngUnMatchCnt(lSiteNum) = 0 Then
                                jdg_color = 4  'Pass:Green
                                If my_MarFlg(lSiteNum) = 0 Then my_MarFlg(lSiteNum) = 1
                                If my_MarFlg(lSiteNum) = 1 Then
                                    If lngUserDelay < my_MinTap(lSiteNum) Then my_MinTap(lSiteNum) = lngUserDelay
                                    If lngUserDelay > my_MaxTap(lSiteNum) Then my_MaxTap(lSiteNum) = lngUserDelay
                                End If
                            Else
                                jdg_color = 38 'No match:Pink
                            End If
                        Else
                            jdg_color = 3 'Alarm occurred:Red
                            If my_MarFlg(lSiteNum) = 1 Then my_MarFlg(lSiteNum) = 2
                        End If
                        If Vod_flg = True Then              'Changes with Ver2.1
                            Worksheets(SheetName & lSiteNum).Cells(rows_step, (lngUserDelay - lStartTap) / lStepTap + 3).Interior.ColorIndex = jdg_color
                        Else
                            Worksheets(SheetName).Cells(rows_step + site_ofst, (lngUserDelay - lStartTap) / lStepTap + 3).Interior.ColorIndex = jdg_color
                            site_cnt = site_cnt + 1
                        End If
                    End If
                Next lSiteNum
        
            Next lngUserDelay
            
            For lSiteNum = 0 To lMaxSite - 1
                If TheExec.sites.site(lSiteNum).Active = True Then
                    TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum) = my_Tap(lSiteNum)
                    my_MidTap(lSiteNum, dataLaneNum + 1) = Int((my_MinTap(lSiteNum) + my_MaxTap(lSiteNum)) / 2)
                End If
            Next lSiteNum
            
        Next vod_val
        
        If Vod_flg = True Then
            Call TheExec.DataManager.GetChanList(sICUL1GPins, -1, chAll, MeasChans(), nChans, nSites, errMsg)
            For lSiteNum = 0 To lMaxSite - 1
                If TheExec.sites.site(lSiteNum).Active = True Then
                    If dataLaneNum = -1 Then
                        lngChArryNum = lSiteNum
                    Else
                        lngChArryNum = (lNumLane - dataLaneNum) * lMaxSite * 2 + lSiteNum
                    End If
                    If ((MeasChans(lngChArryNum) >= 0) And (MeasChans(lngChArryNum) <= 7)) Or _
                       ((MeasChans(lngChArryNum) >= 24) And (MeasChans(lngChArryNum) <= 31)) Or _
                       ((MeasChans(lngChArryNum) >= 48) And (MeasChans(lngChArryNum) <= 55)) Or _
                       ((MeasChans(lngChArryNum) >= 72) And (MeasChans(lngChArryNum) <= 79)) Then
                        lng1gChType = icul1gChannelTypeClkHigh
                    Else
                        lng1gChType = icul1gChannelTypeDataHigh
                    End If
                    Call TheHdw.ICUL1G.Chans(MeasChans(lngChArryNum), lng1gChType).PinLevels.ModifyDiffLevel(icul1gPinLevelVod, my_Vod(lSiteNum, dataLaneNum + 1))
                End If
            Next lSiteNum
        End If
        
    Next dataLaneNum
    '----------------------------------------------------------------------------------------
    
    If (Vod_flg = True) Or (trimflag = 0) Then GoTo End_Seq
    
    
    
    '=== Tap value is adjusted at center, and TAP margin is acquired again. ===
    
    '----------------------------------------------------------------------------------------
    'Tap value of each site and each lane is set to center.
    For dataLaneNum = -1 To lNumLane - 1
        For lSiteNum = 0 To lMaxSite - 1
            If TheExec.sites.site(lSiteNum).Active = True Then
                TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum) = my_MidTap(lSiteNum, dataLaneNum + 1)
            End If
        Next lSiteNum
    Next dataLaneNum
    '----------------------------------------------------------------------------------------
    
    'Initialize Alarm
    With TheHdw.ICUL1G.ImgPins(sICUL1GPins)
        .alarm(icul1gAlarmAll) = icuAlarmRedirect
        .ClearAlarm (icul1gAlarmAll)
    End With
    
'    'Capture setup
'    cpins(0) = sICUL1GPins
'    cDest(0) = sCapPlane
'    lPortNum = 1
'    TheHdw.IDP.SetPMD cDest(0), sCapPad
'    TheHdw.IDP.WritePixel cDest(0), idpColorFlat, 0
    
    '----------------------------------------------------------------------------------------
    skew_channel = 0
    For dataLaneNum = -1 To lNumLane - 1
        rows_step = (dataLaneNum + lNumLane + 7) + (next_ch_cells * skew_channel + vod_max_step * skew_channel) + vod_cnt
    
        Call TheHdw.ICUL1G.ImgPins(sICUL1GPins).GetUserDelayRange(dataLaneNum, nMin, nMax)
        If lngPinStdd = 2 Then
            TheHdw.ICUL1G.ImgPins(sICUL1GPins).serial.EstimateEdgeAlignment 400# * MHz, icul1gClockTypeSerialDDR, lngClockEstDelay(), lngDataEstDelay()
        Else
            ReDim lngClockEstDelay(lMaxSite - 1)
            ReDim lngDataEstDelay(lMaxSite - 1)
            For lSiteNum = 0 To lMaxSite - 1
                lngClockEstDelay(lSiteNum) = 0
                lngDataEstDelay(lSiteNum) = 0
            Next lSiteNum
        End If
        
        site_cnt = 0                                    'Adds with Ver2.1
        For lSiteNum = 0 To lMaxSite - 1
            site_ofst = (lNumLane + 4) * site_cnt * 2   'Adds with Ver2.1
            If TheExec.sites.site(lSiteNum).Active = True Then
                my_Tap(lSiteNum) = TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum)
                If dataLaneNum = -1 Then
                    strVodInfo = "Clock: Vod=" & Round(my_Vod(lSiteNum, dataLaneNum + 1) * 1000, 2) & "[mV]"
                    strTapInfo = "Tap Setting Value = " & my_Tap(lSiteNum) - lngClockEstDelay(lSiteNum)
                Else
                    strVodInfo = "Lane" & dataLaneNum & ": Vod=" & Round(my_Vod(lSiteNum, dataLaneNum + 1) * 1000, 2) & "[mV]"
                    strTapInfo = "Tap Setting Value = " & my_Tap(lSiteNum) - lngDataEstDelay(lSiteNum)
                End If
                Worksheets(SheetName).Cells(rows_step + site_ofst, 1).Value = strVodInfo    'Changes with Ver2.1
                Worksheets(SheetName).Cells(rows_step + site_ofst, 2).Value = strTapInfo
                site_cnt = site_cnt + 1
            End If
        Next lSiteNum
        
        lWidthTap = 100
        column_step = 2
        For lngUserDelay = Int(0 - (lWidthTap / 2)) To Int(0 + (lWidthTap / 2)) Step lStepTap
            DoEvents

            column_step = column_step + 1
            site_cnt = 0                                    'Adds with Ver2.1
            For lSiteNum = 0 To lMaxSite - 1
                site_ofst = (lNumLane + 4) * site_cnt * 2   'Adds with Ver2.1
                If TheExec.sites.site(lSiteNum).Active = True Then
                    If soutaiflag = 1 Then
                        TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum) = TheHdw.ICUL1G.ImgPins(sICUL1GPins).UserDelay(dataLaneNum) + lngUserDelay
                    Else
                        If dataLaneNum = -1 Then
                            If my_Tap(lSiteNum) + lngUserDelay < nMax Then
                                If my_Tap(lSiteNum) + lngUserDelay > nMin Then
                                    TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum) = lngClockEstDelay(lSiteNum) + my_Tap(lSiteNum) + lngUserDelay
                                Else
                                    TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum) = lngClockEstDelay(lSiteNum) + nMin
                                End If
                            Else
                                TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum) = lngClockEstDelay(lSiteNum) + nMax
                            End If
                        Else
                            If my_Tap(lSiteNum) + lngUserDelay < nMax Then
                                If my_Tap(lSiteNum) + lngUserDelay > nMin Then
                                    TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum) = lngDataEstDelay(lSiteNum) + my_Tap(lSiteNum) + lngUserDelay
                                Else
                                    TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum) = lngDataEstDelay(lSiteNum) + nMin
                                End If
                            Else
                                TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum) = lngDataEstDelay(lSiteNum) + nMax
                            End If
                        End If
                    End If
                    Worksheets(SheetName).Cells(lNumLane + 5 + site_ofst, 1).Value = "Site" & lSiteNum       'Changes with Ver2.1
                    Worksheets(SheetName).Cells(lNumLane + 5 + site_ofst, 2).Value = strMipiCaptureMode
                    Worksheets(SheetName).Cells(lNumLane + 5 + site_ofst, column_step).Value = ((lWidthTap / 2) - lWidthTap) + ((column_step - 3) * lStepTap)
                    site_cnt = site_cnt + 1
                End If
            Next lSiteNum
        
                'ADD koga
                cPins(0) = sICUL1GPins
                cDest(0) = sCapPlane.Name
        
            'Start Capturing
            TheHdw.IDP.MultiAcquire cPins, cDest, lPortNum, 1, idpNonAverage, idpAcqNonInterlace, , , idpCurrentPmd
        
            'Wait for the image capture to complete
            TheHdw.IDP.WaitCaptureCompletion
            TheHdw.IDP.WaitTransferCompletion
            
            'Get CaputureResult
            For lSiteNum = 0 To lMaxSite - 1
                If TheExec.sites.site(lSiteNum).Active = True Then
                    TheHdw.IDP.ReadAcquiredFrameCount lSiteNum, cDest(skew_channel), lngCapturedFrames(lSiteNum)
                    'Input AcquireStatus to FlagAry
                    TheHdw.IDP.ReadAcquireStatus lSiteNum, cDest(skew_channel), lngReadAcqFlag(lSiteNum)
                    'Get Capture Alarm Status
                    blAlarmCaptureRelated(lSiteNum) = TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).AlarmStatus(icul1gAlarmCaptureRelated)
                End If
            Next lSiteNum
            
            If (CheckTP = True) And (sNonCompPad <> "") Then           'Adds with Ver2.0
                'Written -1 in no comparison areas
                With TheHdw.IDP
                    .SetPMD sCapPlane, sNonCompPad
                    .SetPMD sExptPlane, sNonCompPad
                    .WritePixel sCapPlane, idpColorFlat, -1
                    .WritePixel sExptPlane, idpColorFlat, -1
                End With
            End If
                    
            site_cnt = 0                                    'Adds with Ver2.1
            For lSiteNum = 0 To lMaxSite - 1
                site_ofst = (lNumLane + 4) * site_cnt * 2   'Adds with Ver2.1
                If TheExec.sites.site(lSiteNum).Active = True Then
                    If CheckTP = True Then
                        'Matching and Jadge
                        With TheHdw.IDP
                            .SetPMD sCapPlane, sCapPad
                            .SetPMD sExptPlane, sCapPad
                            .SetPMD sWorkPlane, sCapPad
                            .WritePixel sWorkPlane, idpColorFlat, 0
                        
                            .LXor sCapPlane, idpColorFlat, sExptPlane, idpColorFlat, sWorkPlane, idpColorFlat
                            .Count sWorkPlane, idpColorFlat, idpCountOutside, 0, 0, idpLimitExclude
                            .ReadResult lSiteNum, sWorkPlane, idpColorFlat, idpReadCount, lngRetMatchVal
                        End With
                        lngUnMatchCnt(lSiteNum) = lngRetMatchVal(idpColorRed)
                    Else
                        lngUnMatchCnt(lSiteNum) = 0
                    End If
                    
                    If (lngReadAcqFlag(lSiteNum) <> idpAcqCompleted) Or (blAlarmCaptureRelated(lSiteNum) = True) Then   'Changes with Ver2.1
                        capFlag(lSiteNum) = 1
                    Else
                        capFlag(lSiteNum) = 0
                    End If
                    
                    If capFlag(lSiteNum) = 0 Then           'Changes with Ver2.1
                        If lngUnMatchCnt(lSiteNum) = 0 Then
                            jdg_color = 4 'Pass:Green
                        Else
                            jdg_color = 38 'No match:Pink
                        End If
                    Else
                        jdg_color = 3 'Alarm occurred:Red
                    End If
                    Worksheets(SheetName).Cells(rows_step + site_ofst, column_step).Interior.ColorIndex = jdg_color
                    site_cnt = site_cnt + 1
                End If
            Next lSiteNum
    
        Next lngUserDelay
        
        For lSiteNum = 0 To lMaxSite - 1
            If TheExec.sites.site(lSiteNum).Active = True Then
                TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).UserDelay(dataLaneNum) = my_Tap(lSiteNum)
            End If
        Next lSiteNum
        
    Next dataLaneNum
    '----------------------------------------------------------------------------------------

End_Seq:
    If my_RunModeFlg = True Then
        TheExec.RunMode = runModeDebug
    End If
    TheHdw.IDP.CaptureTimeOut = my_CapTimeOut

End Function


Public Function ICUL1G_Capture_LpsVt(ByRef sICUL1GPins As String, ByRef sCapPlane As CImgPlane, sCapPad As String, dStartVt As Double, dStopVt As Double, dStepVt As Double, _
                                    Optional CheckTP As Boolean = False, Optional sExptPlane As String = "", Optional sWorkPlane As String = "", Optional sNonCompPad As String = "", Optional sExptFile As String = "") As Long
    
    Dim cPins(0) As String
    Dim cDest(0) As String
    Dim lPortNum As Long
    Dim lng1GBoardNum As Long
    Dim dblSetLpsVt As Double
    
    Dim MeasChans() As Long
    Dim nChans As Long
    Dim nSites As Long
    Dim errMsg As String

    Dim lngRetMatchVal(12) As Long
    Dim lngUnMatchCnt() As Long
    Dim my_LpsVt() As Double
    Dim my_MidLpsVt() As Double
    Dim my_MinLpsVt() As Double
    Dim my_MaxLpsVt() As Double
    Dim my_MarFlg As Long
    Dim my_RunModeFlg As Boolean
    Dim my_CapTimeOut As Double
    
    Dim lSiteNum As Long
    Dim lMaxSite As Long
    Dim capFlag() As Double
    Dim lngCapturedFrames() As Long
    Dim lngReadAcqFlag() As Long
    Dim blAlarmCaptureRelated() As Boolean
    
    Dim SheetName As String
    Dim rows_step As Long
    Dim column_step As Long
    Dim jdg_color As Long                       'Adds with Ver2.1


    lMaxSite = TheExec.sites.ExistingCount
    ReDim my_LpsVt(lMaxSite - 1) As Double
    ReDim my_MidLpsVt(lMaxSite - 1) As Double
    ReDim my_MinLpsVt(lMaxSite - 1) As Double
    ReDim my_MaxLpsVt(lMaxSite - 1) As Double
    ReDim capFlag(lMaxSite - 1)
    ReDim lngCapturedFrames(lMaxSite - 1)
    ReDim lngReadAcqFlag(lMaxSite - 1)
    ReDim blAlarmCaptureRelated(lMaxSite - 1)
    ReDim lngUnMatchCnt(lMaxSite - 1)
    
    If (dStartVt < -1) Or (dStartVt > 2.5) Or (dStopVt < -1) Or (dStopVt > 2.5) Then
        MsgBox "A set value of LPS Vt is illegal." & vbCrLf & _
               "The voltage level can be set within from -1.0V to +2.5V."
        Exit Function
    End If
    
    
    my_RunModeFlg = False
    If TheExec.RunMode = runModeDebug Then
        TheExec.RunMode = runModeProduction
        my_RunModeFlg = True
    End If
    my_CapTimeOut = TheHdw.IDP.CaptureTimeOut
    TheHdw.IDP.CaptureTimeOut = 0.3

    SheetName = "LpsVt"
    Worksheets(SheetName).Cells.Clear
    Worksheets(SheetName).Cells.Interior.ColorIndex = xlNone
    
    'Initialize Alarm
    With TheHdw.ICUL1G.ImgPins(sICUL1GPins)
        .alarm(icul1gAlarmAll) = icuAlarmRedirect
        .ClearAlarm (icul1gAlarmAll)
    End With
    
    'Capture setup
    cPins(0) = sICUL1GPins
    cDest(0) = sCapPlane.Name
    lPortNum = 1
'    TheHdw.IDP.SetPMD cDest(0), sCapPad
    TheHdw.IDP.WritePixel cDest(0), idpColorFlat, 0

    'Load Reference Image
    If CheckTP = True Then
        TheHdw.IDP.SetPMD sExptPlane, sCapPad
        TheHdw.IDP.WritePixel sExptPlane, idpColorFlat, 0
        For lSiteNum = 0 To lMaxSite - 1
            If TheExec.sites.site(lSiteNum).Active Then
                TheExec.Datalog.WriteComment "Loading expect image of site" & lSiteNum & "...  " & sExptFile & " to " & sExptPlane
                TheHdw.IDP.ReadFile lSiteNum, sExptPlane, idpColorFlat, sExptFile, idpFileBinary
            End If
        Next lSiteNum
    End If

    '----------------------------------------------------------------------------------------
    
    For dblSetLpsVt = dStartVt To dStopVt Step dStepVt
        Worksheets(SheetName).Cells(1, (dblSetLpsVt - dStartVt) / dStepVt + 3).Value = dblSetLpsVt
    Next dblSetLpsVt
    
    rows_step = 2
    For lSiteNum = 0 To lMaxSite - 1
        If TheExec.sites.site(lSiteNum).Active = True Then
            Call TheExec.DataManager.GetChanList(sICUL1GPins, lSiteNum, chAll, MeasChans(), nChans, nSites, errMsg)
            If (MeasChans(0) >= 0) And (MeasChans(0) <= 23) Then
                lng1GBoardNum = 17
            ElseIf (MeasChans(0) >= 24) And (MeasChans(0) <= 47) Then
                lng1GBoardNum = 19
            ElseIf (MeasChans(0) >= 48) And (MeasChans(0) <= 71) Then
                lng1GBoardNum = 16
            ElseIf (MeasChans(0) >= 72) And (MeasChans(0) <= 95) Then
                lng1GBoardNum = 20
            Else
                MsgBox "Error : As for the specified pin, it is not ICUL1G pin or wrong Ch is set."
                lng1GBoardNum = 17
                Exit For
            End If
            
            Worksheets(SheetName).Cells(rows_step, 1).Value = "Site" & lSiteNum
            Worksheets(SheetName).Cells(rows_step, 2).Value = "ICUL1G Slot Number = " & lng1GBoardNum
            
            my_LpsVt(lSiteNum) = TheHdw.ICUL1G.board(lng1GBoardNum).ReadMipiLpThreshold
            my_MinLpsVt(lSiteNum) = dStopVt
            my_MaxLpsVt(lSiteNum) = dStartVt
            my_MarFlg = 0
                
            For dblSetLpsVt = dStartVt To dStopVt Step dStepVt
                'Set LPS LpFullStateSync Mode
                TheHdw.ICUL1G.MipiCaptureMode = icul1gMipiLpFullStateSync
'                TheHdw.ICUL1G.MipiCaptureMode = icul1gMipiLp11StateSync
                'Set LPS VThreshold Level
                TheHdw.ICUL1G.board(lng1GBoardNum).ModifyMipiLpThreshold dblSetLpsVt
                TheHdw.WAIT 0.01
                
                'ADD koga
                cPins(0) = sICUL1GPins
                cDest(0) = sCapPlane.Name
                
                'Start Capturing
                TheHdw.IDP.MultiAcquire cPins(), cDest(), lPortNum, 1, idpNonAverage, idpAcqNonInterlace, , lSiteNum, idpCurrentPmd
    
                'Wait for the image capture to complete
                TheHdw.IDP.WaitCaptureCompletion
                TheHdw.IDP.WaitTransferCompletion
                
                'Get CaputureResult
                TheHdw.IDP.ReadAcquiredFrameCount lSiteNum, cDest(0), lngCapturedFrames(lSiteNum)
                'Input AcquireStatus to FlagAry
                TheHdw.IDP.ReadAcquireStatus lSiteNum, cDest(0), lngReadAcqFlag(lSiteNum)
                'Get Capture Alarm Status
                blAlarmCaptureRelated(lSiteNum) = TheHdw.ICUL1G.ImgPins(sICUL1GPins, lSiteNum).AlarmStatus(icul1gAlarmCaptureRelated)
                
                If (CheckTP = True) And (sNonCompPad <> "") Then           'Adds with Ver2.0
                    'Written -1 in no comparison areas
                    With TheHdw.IDP
                        .SetPMD sCapPlane, sNonCompPad
                        .SetPMD sExptPlane, sNonCompPad
                        .WritePixel sCapPlane, idpColorFlat, -1
                        .WritePixel sExptPlane, idpColorFlat, -1
                    End With
                End If
                
                If CheckTP = True Then
                    'Matching and Jadge
                    With TheHdw.IDP
                        .SetPMD sCapPlane, sCapPad
                        .SetPMD sExptPlane, sCapPad
                        .SetPMD sWorkPlane, sCapPad
                        .WritePixel sWorkPlane, idpColorFlat, 0
                    
                        .LXor sCapPlane, idpColorFlat, sExptPlane, idpColorFlat, sWorkPlane, idpColorFlat
                        .Count sWorkPlane, idpColorFlat, idpCountOutside, 0, 0, idpLimitExclude
                        .ReadResult lSiteNum, sWorkPlane, idpColorFlat, idpReadCount, lngRetMatchVal
                    End With
                    lngUnMatchCnt(lSiteNum) = lngRetMatchVal(idpColorRed)
                Else
                    lngUnMatchCnt(lSiteNum) = 0
                End If
                
                If (lngReadAcqFlag(lSiteNum) <> idpAcqCompleted) Or (blAlarmCaptureRelated(lSiteNum) = True) Then
                    capFlag(lSiteNum) = 1
                Else
                    capFlag(lSiteNum) = 0
                End If
                
                If capFlag(lSiteNum) = 0 Then           'Changes with Ver2.1
                    If lngUnMatchCnt(lSiteNum) = 0 Then
                        jdg_color = 4 'Pass:Green
                        If my_MarFlg = 0 Then my_MarFlg = 1
                        If my_MarFlg = 1 Then
                            If dblSetLpsVt < my_MinLpsVt(lSiteNum) Then my_MinLpsVt(lSiteNum) = dblSetLpsVt
                            If dblSetLpsVt > my_MaxLpsVt(lSiteNum) Then my_MaxLpsVt(lSiteNum) = dblSetLpsVt
                        End If
                    Else
                        jdg_color = 38 'No match:Pink
                    End If
                Else
                    jdg_color = 3 'Alarm occurred:Red
                    If my_MarFlg = 1 Then my_MarFlg = 2
                End If
                Worksheets(SheetName).Cells(rows_step, (dblSetLpsVt - dStartVt) / dStepVt + 3).Interior.ColorIndex = jdg_color
            Next dblSetLpsVt
            
            TheHdw.ICUL1G.board(lng1GBoardNum).ModifyMipiLpThreshold my_LpsVt(lSiteNum)
            If my_MarFlg = 0 Then
                my_MidLpsVt(lSiteNum) = 9999
            Else
                my_MidLpsVt(lSiteNum) = (my_MinLpsVt(lSiteNum) + my_MaxLpsVt(lSiteNum)) / 2
            End If
            Worksheets(SheetName).Cells(rows_step, (dblSetLpsVt - dStartVt) / dStepVt + 3).Value = "Center LPS Vt = " & my_MidLpsVt(lSiteNum)
            rows_step = rows_step + 1
            
        End If
    Next lSiteNum
    '----------------------------------------------------------------------------------------
    
    If my_RunModeFlg = True Then
        TheExec.RunMode = runModeDebug
    End If
    TheHdw.IDP.CaptureTimeOut = my_CapTimeOut

End Function



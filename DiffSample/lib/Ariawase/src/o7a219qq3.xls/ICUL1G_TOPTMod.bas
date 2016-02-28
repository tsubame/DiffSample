Attribute VB_Name = "ICUL1G_TOPTMod"
Option Explicit

Private Const IsCaputreTimeComment As Boolean = False
Private Const IsToptComment As Boolean = False

Public Flg_Acq As Boolean
Public Flg_Acq_Site(nSite) As Boolean
Public Flg_Active_Chip(nSite) As Boolean

Private keyName As String
Private capPins As String
Private AcqCheckCounter As Double
Private PlaneGetCheckCounter As Double
Private captureTimer As Double
Private CaptureTimeOut As Double
Public Dout As EnumDOUT

'FREQ CONST
Public Const lngFreqSampleNum As Long = 1000
Public Const Freq_lim As Long = 50

Private Const EEE_STANDBY = 10

' Ver1.1 2013/01/30 H.Arikawa CaptureInstance名を特定する為、split処理を追加。

Private Const ARG_CAPTURE_ERROR_INSTANCE_NAME       As Long = 0
Private Const ARG_IMAGE_ACQUIRE_INSTANCE_NAME       As Long = 1
Private Const ARG_ACQUIRE_PLANE_SOURCE_ZONE         As Long = 2
Private Const ARG_PROCESS_PLANE_NAME                As Long = 3
Private Const ARG_PROCESS_PLANE_TARGET_ZONE         As Long = 4
Private Const ARG_COPY_PLANE_MODE_NAME              As Long = 5
Private Const ARG_SYNC_CODE_JUDGE_MODE_NAME         As Long = 6
Private Const ARG_SYNC_CODE_LEFT_REF_NAMELIST       As Long = 7
Private Const ARG_SYNC_CODE_LEFT_JUDGE_ZONELIST     As Long = 8
Private Const ARG_SYNC_CODE_LEFT_JUDGE_MODE_NAME    As Long = 9
Private Const ARG_SYNC_CODE_RIGHT_REF_NAMELIST      As Long = 10
Private Const ARG_SYNC_CODE_RIGHT_JUDGE_ZONELIST    As Long = 11
Private Const ARG_SYNC_CODE_RIGHT_JUDGE_MODE_NAME   As Long = 12
Private Const ARG_SYNC_CODE_TOP_REF_NAMELIST        As Long = 13
Private Const ARG_SYNC_CODE_TOP_JUDGE_ZONELIST      As Long = 14
Private Const ARG_SYNC_CODE_TOP_JUDGE_MODE_NAME     As Long = 15
Private Const ARG_SYNC_CODE_BOTTOM_REF_NAMELIST     As Long = 16
Private Const ARG_SYNC_CODE_BOTTOM_JUDGE_ZONELIST   As Long = 17
Private Const ARG_SYNC_CODE_BOTTOM_JUDGE_MODE_NAME  As Long = 18

Public Function FW_Acquire_1G(ByRef acqParam As CAcquireFunctionInfo, _
                                ByRef pParamPlane As CParamPlane) As Long

    On Error GoTo ErrHandler

    Dim ICUL1G_Pin(0) As String
    Dim ICUL1G_Plane(0) As String

    Dim site As Long

    '++++ TEST CONDITION ++++++++++++++++++++++++++++++++++
    keyName = acqParam.Arg(0)
    
    '++++ CAPTURE AVERAGE +++++++++++++++++++++++++++++++++
    Dim averageNum As Long
    averageNum = GetCaptureParamAverageCount(acqParam.Arg(0))
    
    Dim averageMode As IdpAverageMode
    averageMode = GetCaptureParamAverageMode(acqParam.Arg(0))

    If TheVarBank.IsExist("PlaneGetCheckCounter") = False Then
        Call TheVarBank.Add("PlaneGetCheckCounter", 0, False, "PlaneGetCheckCounter")
        PlaneGetCheckCounter = TheVarBank.Value("PlaneGetCheckCounter")
    End If

    If Flg_Simulator = 1 Then Exit Function
    
    Dim acqPlane As CImgPlane
    Set acqPlane = pParamPlane.plane
    If acqPlane Is Nothing Then
        TheHdw.TOPT.Recall
        TheExec.Datalog.WriteComment "Recall ACQUIRE:" & keyName & " Recall Count:" & PlaneGetCheckCounter
        PlaneGetCheckCounter = PlaneGetCheckCounter + 1
        If PlaneGetCheckCounter >= 20 Then
            Call WaitSet(200 * mS)
        Else
            Call WaitSet(10 * mS)
        End If
        FW_Acquire_1G = EEE_STANDBY
        Exit Function
    End If
        
    '++++ Capture PinGroup +++++++++++++++++++++++++++++
    capPins = acqParam.Arg(9)

    '++++ ERASE +++++++++++++++++++++++++++++++++++++++++++
    AcqCheckCounter = 0
        
    '++++ TIME OUT SET ++++++++++++++++++++++++++++++++++++
    CaptureTimeOut = GetCaptureTimeOut(averageNum)
    Call CaptureTimeOutStart

    '++++ CAPTURE +++++++++++++++++++++++++++++++++++++++++
    '''TheHdw.ICUL1G.Pins(capPins).Connect              'ICUL1G 不具合対策
    ICUL1G_Pin(0) = capPins
    ICUL1G_Plane(0) = acqPlane.Name
    
'@@@@@@@ Get ICUL1G SKEW @@@@@@@@@@@@@
'    Call ICUL1G_Capture_LpsVt(capPins, acqPlane, acqPlane.CurrentPmdName, 0.2, 0.5, 0.01)
'    Call ICUL1G_Capture_Skew(capPins, acqPlane, acqPlane.CurrentPmdName, 125, 285, 5, , , , , , True, 0.08, -0.09, -0.01)
'    Call ICUL1G_Capture_Skew(capPins, acqPlane, acqPlane.CurrentPmdName, 125, 285, 5, , , , , , False, 0.08, -0.09, -0.01)
    
    'FreqCntStart+++++++++++++++++++++++++++++++
    Call FreqCount_Condi(capPins)
    'FreqCntEnd+++++++++++++++++++++++++++++++++
        
        TheHdw.WAIT 20 * mS

    '===== IMAGE DATA ACQUIRE (HSCIS ACQUIRE ONLY) ========
    Call acqPlane.SetPMD(acqPlane.BasePMD.Name)
    TheHdw.IDP.MultiAcquire ICUL1G_Pin, ICUL1G_Plane, 1, averageNum, averageMode, idpAcqNonInterlace  '2013/02/04
    
    Flg_Acq = True
    Call GetFlagAcqSite                                                                               '2013/02/04
    
        'TheHdw.IDP.CaptureTimeoutで設定されるIDP側のWait Timeと、TheHdw.TOPT.WaitCondition
        'で設定される1 Acquire Instance実行完了までの待ち時間は、以下の関係を満たさねばならない。
        '　　CaptureTimeout(IDP) < WaitCondition(TOPT)
        'The values of properties "TheHdw.IDP.CaptureTimeout" and "TheHdw.TOPT.WaitCondition"
        'must satisfy the relation, aobve. (Based on a trouble analysis of IMX164).
        Call TheHdw.TOPT.WaitCondition(toptAcquire, CaptureTimeOut * 1000 * averageNum * 2 + 10000)

    If IsToptComment = True Then TheExec.Datalog.WriteComment "TOPT: ACQUIRE " & keyName & " -> " & acqPlane.Name
    '画像キャプチャが成功したことをフレームワークへ報告する
    FW_Acquire_1G = TL_SUCCESS
    Exit Function
ErrHandler:
    '実行を止める場合はフレームワークにTL_ERRORを返す
    FW_Acquire_1G = TL_ERROR
InputErr:
    MsgBox "Input Error! @Image ACQTBL Sheet! " & keyName & " : " & "Msg"  '2013/02/04

End Function

Public Function FW_PostAcquire_1G(ByRef acqParam As CAcquireFunctionInfo, _
                                ByRef pParamPlane As CParamPlane) As Long

    On Error GoTo ErrHandler

    Dim CaptureErr(nSite) As Double
    Dim blnAlarm As Boolean
    Dim site As Long
    Dim FreqCount(nSite) As Double
    Dim dblFreq(nSite) As Double
    Dim lngFreqCnt(nSite) As Long
    
    PlaneGetCheckCounter = TheVarBank.Value("PlaneGetCheckCounter")
    
    '------------------------------------------------------
    If Flg_Simulator = 1 Then
        For site = 0 To nSite
            CaptureErr(site) = 0
            FreqCount(site) = 0
        Next site
        TheResult.Add keyName, CaptureErr
        TheResult.Add "FreqErr_" & keyName, FreqCount
        Exit Function
    End If
    '------------------------------------------------------

    '++++ Capture PinGroup +++++++++++++++++++++++++++++
    capPins = acqParam.Arg(9)
    
    Call GetFlagActive
    Call AcqSiteActive

    '----- ICUD Capture WAIT &  ERROR CODE CHECK -------------------
    TheHdw.IDP.WaitCaptureCompletion
    TheHdw.IDP.WaitTransferCompletion
    
    ' ============= Freq Count =============================================
    TheHdw.ICUL1G.ImgPins(capPins).FreqCtr.ReadCounter lngFreqCnt
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            dblFreq(site) = lngFreqCnt(site) / (TheHdw.ICUL1G.FreqCtrSampleLengthUnitPeriod * lngFreqSampleNum) / 1000000
            If Flg_Debug = 1 Then
                TheExec.Datalog.WriteComment keyName & ": Freq(" & site & ") = " & dblFreq(site) & " [Mhz]"
            End If
            Call FreqCnt(CapFreq_Typ - Freq_lim, CapFreq_Typ + Freq_lim, dblFreq(site), FreqCount(site))
        End If
    Next site
    ' ============= Freq Count End =========================================
    
    '++++ ICUL1G Error Check +++++++++++++++
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            CaptureErr(site) = 0
            blnAlarm = TheHdw.ICUL1G.ImgPins(capPins, site).AlarmStatus(icul1gAlarmCaptureRelated)
            If blnAlarm = True Then
                CaptureErr(site) = CaptureErr(site) + 10
                If Flg_Debug = 1 Then
                    TheExec.Datalog.WriteComment keyName & ": CaptureErr(" & CStr(site) & ") = " & CStr(CaptureErr(site))
                End If
            End If
        End If
    Next site
    
    TheResult.Add keyName, CaptureErr
    TheResult.Add "FreqErr_" & keyName, FreqCount

    Flg_Acq = False
    
    Call ReturnActiveSite
    If IsToptComment = True Then TheExec.Datalog.WriteComment "TOPT: COMPLET " & keyName

    '画像トランスファが成功したことをフレームワークへ報告する
    FW_PostAcquire_1G = TL_SUCCESS
    Exit Function
ErrHandler:
    '実行を止める場合はフレームワークにTL_ERRORを返す
    FW_PostAcquire_1G = TL_ERROR

End Function

'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'MMMMMMMMMMMMMMMMMMMMMMMMMMMM FREQ_CNT MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
Public Sub FreqCount_Condi(ByVal mipi_lane As String)

    With TheHdw.ICUL1G.ImgPins(mipi_lane).FreqCtr
        .Clear
        .SampleLength = lngFreqSampleNum
        .PrepareStart
        .StartBySyncCodeOnMultiAcquire
    End With
            
End Sub

Private Sub CaptureTimeOutStart()
    '--- Time Measure  START ! --------
    captureTimer = TheExec.timer(0)
    '----------------------------------
End Sub

Private Sub WaitSet(ByVal waitTime As Double)
    If TheExec.RunOptions.AutoAcquire = True Then
        Call TheHdw.TOPT.WAIT(toptTimer, waitTime * 1000)
    Else
        Call TheHdw.WAIT(waitTime)
    End If
End Sub

Public Sub GetFlagActive()
    Dim site As Long
    For site = 0 To nSite
        Flg_Active_Chip(site) = TheExec.sites.site(site).Active
    Next site
End Sub
Public Sub GetFlagAcqSite()
    Dim site As Long
    For site = 0 To nSite
        Flg_Acq_Site(site) = TheExec.sites.site(site).Active
    Next site
End Sub
Public Sub AcqSiteActive()
    Dim site As Long
    For site = 0 To nSite
        TheExec.sites.site(site).Active = Flg_Acq_Site(site)
    Next site
End Sub

Public Function GetCaptureTimeOut(ByVal averageNum As Long) As Double
    If Flg_Simulator = 0 Then
        GetCaptureTimeOut = (TheIDP.CaptureTimeOut) * averageNum
    Else
        GetCaptureTimeOut = (0.5 * S) * averageNum
    End If
End Function

Public Sub FreqCnt(ByVal Lo_Lim As Double, ByVal Hi_Lim As Double, ByVal indata As Double, OutData As Double)   '2013/02/04
    
    If indata < Lo_Lim Or Hi_Lim < indata Then
        OutData = OutData + 1
    End If

End Sub
Public Sub ReturnActiveSite()  '2013/02/04
    Dim site As Long
    For site = 0 To nSite
        TheExec.sites.site(site).Active = Flg_Active_Chip(site)
    Next site
End Sub


Public Function PutImageInto_Common() As Long

    'Test Instanceシートからパラメータを取得。(To get "PutImageInto" parameters from "Test Instances" worksheet)
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, -1) Then
        Err.Raise 9999, "PutImageInto_Common", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
    End If
    
    'Image ACQTBLにおいて画像取得を行ったInstance名を取得(To get image acquire instance name on "Image ACQTBL" worksheet)
    Dim acqInstanceName As String
    acqInstanceName = ArgArr(ARG_IMAGE_ACQUIRE_INSTANCE_NAME)
    
    '=== 2013/01/30 Add H.Arikawa start===
    Dim tmpInstanceName() As String
    tmpInstanceName = Split(acqInstanceName, ",")
    acqInstanceName = tmpInstanceName(UBound(tmpInstanceName))
    '=== 2013/01/30 Add H.Arikawa end===
    
    Dim thisCaptureError() As Double
    Dim FreqErr_tmp() As Double
    With TheResult
        .GetResult acqInstanceName, thisCaptureError
        .Delete acqInstanceName
        .GetResult "FreqErr_" & acqInstanceName, FreqErr_tmp
        .Delete "FreqErr_" & acqInstanceName
    End With

    'Parameter Bankに登録されている取り込み画像情報を取得(To get parameter bank item)
    Dim acqParam As CParamPlane
    Set acqParam = TheParameterBank.Item(acqInstanceName)

    '取り込みハードウェア毎特有の取り込み判定処理にて、キャプチャエラー情報を更新
    '   (To refresh capture error result with a procedure specific to each capture system hardware)
    Call HardwareAcquireStatusCheck(thisCaptureError, FreqErr_tmp, acqParam)
    
    'キャプチャエラー項目用の結果を登録
    '   (To register capture error result to the Result Manager)
    Dim captureErrorTestItem As String
    captureErrorTestItem = ArgArr(ARG_CAPTURE_ERROR_INSTANCE_NAME)
''    TheResult.Add captureErrorTestItem, thisCaptureError

If TheResult.IsExist(captureErrorTestItem) = True Then
    Dim tmp_1() As Double
    Dim tmp_2(nSite) As Double

    TheResult.GetResult captureErrorTestItem, tmp_1
    Call GetSum(tmp_2, tmp_1, thisCaptureError)
    TheResult.Delete (captureErrorTestItem)
    TheResult.Add captureErrorTestItem, tmp_2
Else
    TheResult.Add captureErrorTestItem, thisCaptureError
End If

    '取り込み画像を、画像処理用プレーンにコピー
    '   (To copy acquire image to a plane for image process)
    Dim rawPlane As CImgPlane
    Call GetFreePlane(rawPlane, ArgArr(ARG_PROCESS_PLANE_NAME), acqParam.plane.BitDepth, , acqInstanceName)
    Call Copy(acqParam.plane, ArgArr(ARG_ACQUIRE_PLANE_SOURCE_ZONE), EEE_COLOR_ALL, rawPlane, ArgArr(ARG_PROCESS_PLANE_TARGET_ZONE), EEE_COLOR_ALL)

    Set acqParam.plane = rawPlane

End Function

Public Function HardwareAcquireStatusCheck(ByRef captureErrorInfo() As Double, ByRef FreqErrorInfo() As Double, ByRef acqParam As CParamPlane)
    
    Dim site As Long
    Dim rtnAcquireStatus As Long
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            TheHdw.IDP.ReadAcquireStatus site, acqParam.plane.Name, rtnAcquireStatus
            If rtnAcquireStatus = idpAcqCompleted Then                                                  '2013/02/21 修正
                captureErrorInfo(site) = captureErrorInfo(site) + FreqErrorInfo(site)
            Else
                captureErrorInfo(site) = captureErrorInfo(site) + FreqErrorInfo(site) + 100             'CAPTURE NG
            End If
        End If
    Next site
    
End Function

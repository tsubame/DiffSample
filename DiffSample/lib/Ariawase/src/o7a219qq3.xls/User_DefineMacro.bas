Attribute VB_Name = "User_DefineMacro"
Option Explicit

Public Const gDefaultCaptureDir As String = "Z:\imx145\"

'*************************************************
'**     For Mode Set                            **
'*************************************************
Public Enum EnumModeSet
    Futei_1 = -1
    default = 0         'Init Def '@preStandby_f,Ofte,frmpdl_f,cnvte_f,dsgte_f
    SUN = 1             'Def
    SUN_TEMP = 2        'IDLVRL,DK14TE  DK16TE
    OTP = 3
    SCAN = 4
    LOGIC = 5
End Enum

Public Enum EnumModeSetVcp
    Futei_2 = -1
    FROM_VCP = 0        'From VCP
    FROM_TESTER = 1     'From TESTER
End Enum
'*************************************************

Public PREVIOUS_CONDITION As EnumModeSet
Public PREVIOUS_BIAS_CONDITION As EnumModeSetVcp
Public PREVIOUS_Index As String
Public PREVIOUS_FPGA_PINS As Double
Public PREVIOUS_CLOCK_PINS As Double

Public Flg_Sync_Cap As Boolean

Private Const EEE_AUTO_TYPE_NUMBER As Long = 145
Private Const EEE_AUTO_OFFSET_FILE_PATH As String = "..\parameter_pq\BPC"

'概要:
'   自動化管轄外の関数群
'
'目的:
'　 ユーザにやってもらいたい関数はここに記述する
'
'作成者:
'   2012/02/03 Ver0.1 D.Maruyama
'   2012/10/18 Ver0.2 H.Arikawa
'Std_SetUpでコールされる
Public Sub TypeCustomFlagSet()

    '/* === SETUP DEBUG FLAGS ========================================= */
    Flg_Scrn = 0                                'SCRN FLAG                       1:ON  0:OFF  2:Low scrn only
'    Sw_Cbar = 0                                'DISPLAY COLOR BAR CHAT          1:ON   0:OFF
    Flg_Print = 0                               'DATALOG ARRANGE FLAG            1:ON   0:OFF
'    Flg_Shmoo = 0                              'Shmoo FLAG                      1:ON  0:OFF  2:Low scrn only
    Flg_DacLog = 0                              'DAC LOG FLAG                     1:ON  0:OFF
    Flg_Capture = 0                             'Capture Data Output FLAG       3:DK(1V,8V,15V)(Sync-Code) 2:Mura(Sync-Code) 1:Mura(HL-DK)  0:OFF
    Flg_Cnd = 0                                 'Create ConditionCheck          1:OFF   0:ON
'    Flg_OTP_BLOW = 0                            '1:Exex Blow.  other: Don't blow. for ChipID エラーとなったのでとりあえずコメントアウト→Miike
'    Flg_OTP_BLOW_SHINRAISEI = 0                 'OTPBLOW(SHINRAISEI no tokiha Flg_OTP_BLOW to Flg_OTP_BLOW_SHINRAISEI wo "1")
    Flg_Sync_Cap = False                        'For SyncCode Output Debug FLAG  True:ON False:OFF
    
    If Flg_Tenken = 0 And Flg_AutoMode = True Then
        TheExec.RunOptions.AutoAcquire = True   'TOPT Enable FLAG
        Flg_Scrn = 1                            'SCRN FLAG      1:ON  0:OFF  2:Low scrn only
'        Flg_OTP_BLOW = 1                        '1:Exex Blow.  other: Don't blow. for ChipID エラーとなったのでとりあえずコメントアウト→Miike
    End If

End Sub

'JobInterFaceでコールされる
Public Function GetTypeNumber() As Long

    GetTypeNumber = EEE_AUTO_TYPE_NUMBER

End Function

'JobInterFaceでコールされる
Public Sub ModifyChipAddressArray( _
    ByRef aryXAdr() As Long, ByRef aryYAdr() As Long, _
    ByVal lxAdr As Long, ByVal lYAdr As Long)

    Dim site As Long
    For site = 0 To nSite
        aryXAdr(site) = lxAdr - site
        aryYAdr(site) = lYAdr + site
    Next site

End Sub

'ShirotenCheck MarginCheck SiteCheckでコールされる
Public Function ActiveSiteCheck(ByVal site As Long) As Boolean
    
    '白点コブとマージン測定対象chipを選定
    'DC,Logic,画が出てないChipは、測定しない。
'''    If (LastBin(site) >= 25) Or (LastBin(site) = 14) Or (LastBin(site) = 12) Then
    If (LastBin(site) >= 24) Or (LastBin(site) = 14) Or (LastBin(site) = 12) Then
        ActiveSiteCheck = False
    Else
        ActiveSiteCheck = True
    End If
    
End Function
'GetCSVからコールされる
Public Function GetCSVFilePath() As String

    GetCSVFilePath = EEE_AUTO_OFFSET_FILE_PATH

End Function

'Qknee特性を取得する。(汎用ラップ関数 暫定版)
'必ずOFの処理内で関数をCallする事。
Public Sub GetQknee(ByVal FunctionName As String, Optional ByVal start_lux As Double = 0, Optional ByVal step_lux As Double = 5, Optional ByVal cntLP As Long = 30, Optional ByVal ColorMap As String = "Bayer2x4")

    Dim site As Long
    Dim ErrInfo(nSite) As Double
    Dim i As Long
    '========== ExcelSheet Write ==============================================================
    ' *** For worksheet ***
    Dim LuxChk_wkst As Worksheet
    
    TheExec.RunMode = runModeProduction     '!!!For Error window
    
    ' *** Worksheet search & create if not present ***
    For Each LuxChk_wkst In ActiveWorkbook.Sheets
        If LuxChk_wkst.Name = "Lux vs Sens" Then
            GoTo LuxChk_sht
        End If
    Next
    
    ' *** Create new worksheet ***
    Set LuxChk_wkst = ActiveWorkbook.Sheets.Add()
    ' *** Set worksheet name ***
    LuxChk_wkst.Name = "Lux vs Sens"
    
LuxChk_sht:
    LuxChk_wkst.Select

    ' *** Write item name ***
    Cells(1, 2) = "@" & NormalJobName & " Lux vs Sens"
    Cells(2, 2) = "OPT[Lux]"
    Select Case ColorMap
        Case "Bayer2x4"
            Dim Hl_senr1() As Double
            Dim Hl_sengr1() As Double
            Dim Hl_sengb1() As Double
            Dim Hl_senb1() As Double
            Dim Hl_senr2() As Double
            Dim Hl_sengr2() As Double
            Dim Hl_sengb2() As Double
            Dim Hl_senb2() As Double
            For site = 0 To nSite
                Cells(2, 3 + site * 8) = "HL_SENR1_" & site & "[mV]"
                Cells(2, 4 + site * 8) = "HL_SENGR1_" & site & "[mV]"
                Cells(2, 5 + site * 8) = "HL_SENGB1_" & site & "[mV]"
                Cells(2, 6 + site * 8) = "HL_SENB1_" & site & "[mV]"
                Cells(2, 7 + site * 8) = "HL_SENR2_" & site & "[mV]"
                Cells(2, 8 + site * 8) = "HL_SENGR2_" & site & "[mV]"
                Cells(2, 9 + site * 8) = "HL_SENGB2_" & site & "[mV]"
                Cells(2, 10 + site * 8) = "HL_SENB2_" & site & "[mV]"
            Next site
        Case "Bayer2x2"
            Dim Hl_senr() As Double
            Dim Hl_sengr() As Double
            Dim Hl_sengb() As Double
            Dim Hl_senb() As Double
            For site = 0 To nSite
                Cells(2, 3 + site * 8) = "HL_SENR_" & site & "[mV]"
                Cells(2, 4 + site * 8) = "HL_SENGR_" & site & "[mV]"
                Cells(2, 5 + site * 8) = "HL_SENGB_" & site & "[mV]"
                Cells(2, 6 + site * 8) = "HL_SENB_" & site & "[mV]"
            Next site
        Case Else
    End Select
    
    '===== Get Acquire Instance Name =====
    Dim tmpFunctionName() As String
    Dim AcqFunctionName As String
        
    tmpFunctionName = Split(FunctionName, "_Con")
    AcqFunctionName = tmpFunctionName(0) & "_Acq1"
    
    If TheParameterBank.IsExist(AcqFunctionName) Then Call TheParameterBank.Delete(AcqFunctionName)
    
    '======= OptLux Value Setting =====
    For i = 0 To cntLP - 1 Step 1
        '======= ConditionSetting =======
        Call TheCondition.SetCondition(FunctionName)
        If start_lux = 0 Then
            Call OptSet("DARK")
        Else
            Call NSIS_II.SetDevices(Level:=start_lux, Shutter:=0)
            Call OptStatus
        End If
        TheHdw.WAIT 500 * mS
        
        '++++ CAPTURE AVERAGE +++++++++++++++++++++++++++++++++
        Call TheImageTest.RetryAcquire(AcqFunctionName, "FWImageAcquire")
        Call TheImageTest.RetryAcquire(AcqFunctionName, "FWPostImageAcquire")
        Call PutImageInto_Common
    ' #### HL_SEN ####
        '=== ここにユーザーがHL感度を求める処理を貼り付ける。 ===
        Call HL_Process_For_Qknee(AcqFunctionName)
        
        If TheParameterBank.IsExist(AcqFunctionName) Then Call TheParameterBank.Delete(AcqFunctionName)
        
        'Output WorkSheet
        If start_lux = 0 Then
            Cells(i + 3, 2) = start_lux
        Else
            Cells(i + 3, 2) = NSIS_II.Level
        End If
        
        Select Case ColorMap
            Case "Bayer2x4"
                TheResult.GetResult "HL_SENR1", Hl_senr1
                TheResult.GetResult "HL_SENGR1", Hl_sengr1
                TheResult.GetResult "HL_SENGB1", Hl_sengb1
                TheResult.GetResult "HL_SENB1", Hl_senb1
                TheResult.GetResult "HL_SENR2", Hl_senr2
                TheResult.GetResult "HL_SENGR2", Hl_sengr2
                TheResult.GetResult "HL_SENGB2", Hl_sengb2
                TheResult.GetResult "HL_SENB2", Hl_senb2
                TheResult.Delete "HL_SENR1"
                TheResult.Delete "HL_SENGR1"
                TheResult.Delete "HL_SENGB1"
                TheResult.Delete "HL_SENB1"
                TheResult.Delete "HL_SENR2"
                TheResult.Delete "HL_SENGR2"
                TheResult.Delete "HL_SENGB2"
                TheResult.Delete "HL_SENB2"
                For site = 0 To nSite
                    If TheExec.sites.site(site).Active = True Then
                        Cells(i + 3, 3 + site * 8) = Hl_senr1(site) * 1000
                        Cells(i + 3, 4 + site * 8) = Hl_sengr1(site) * 1000
                        Cells(i + 3, 5 + site * 8) = Hl_sengb1(site) * 1000
                        Cells(i + 3, 6 + site * 8) = Hl_senb1(site) * 1000
                        Cells(i + 3, 7 + site * 8) = Hl_senr2(site) * 1000
                        Cells(i + 3, 8 + site * 8) = Hl_sengr2(site) * 1000
                        Cells(i + 3, 9 + site * 8) = Hl_sengb2(site) * 1000
                        Cells(i + 3, 10 + site * 8) = Hl_senb2(site) * 1000
                    End If
                Next site
            Case "Bayer2x2"
                TheResult.GetResult "HL_SENR", Hl_senr
                TheResult.GetResult "HL_SENGR", Hl_sengr
                TheResult.GetResult "HL_SENGB", Hl_sengb
                TheResult.GetResult "HL_SENB", Hl_senb
                TheResult.Delete "HL_SENR"
                TheResult.Delete "HL_SENGR"
                TheResult.Delete "HL_SENGB"
                TheResult.Delete "HL_SENB"
                For site = 0 To nSite
                    If TheExec.sites.site(site).Active = True Then
                        Cells(i + 3, 3 + site * 8) = Hl_senr(site) * 1000
                        Cells(i + 3, 4 + site * 8) = Hl_sengr(site) * 1000
                        Cells(i + 3, 5 + site * 8) = Hl_sengb(site) * 1000
                        Cells(i + 3, 6 + site * 8) = Hl_senb(site) * 1000
                    End If
                Next site
            Case Else
        End Select
        
        start_lux = start_lux + step_lux
    Next i
        
End Sub

'=== ユーザーにカスタム(Copy&Paste&簡易変更)してもらう関数 ===
'=== ①タイプ毎の感度算出処理を貼り付ける ===
'=== ②所定の部分をacqInstanceNameに変更する ===

Public Sub HL_Process_For_Qknee(ByVal acqInstanceName As String)

'=== ↓ここに貼り付けて下さい。 ===
'=== [参考]IMX175での例 (使用時は削除して下さい!!)===
' #### HL_SEN ####

        Dim site As Long

        Dim HL_ERR_Param As CParamPlane
        Dim HL_ERR_DevInfo As CDeviceConfigInfo
        Dim HL_ERR_Plane As CImgPlane
        Set HL_ERR_Param = TheParameterBank.Item(acqInstanceName)  '←acqInstanceNameに変更して下さい。
        Call TheIDP.PlaneBank.Delete(acqInstanceName)              '←acqInstanceNameに変更して下さい。
        Set HL_ERR_DevInfo = HL_ERR_Param.DeviceConfigInfo
        Set HL_ERR_Plane = HL_ERR_Param.plane

        Dim HL_ERR_LSB() As Double
        HL_ERR_LSB = HL_ERR_DevInfo.Lsb.AsDouble

        'OPBクランプ
        Dim sPlane1 As CImgPlane
        Call GetFreePlane(sPlane1, HL_ERR_Plane.planeGroup, HL_ERR_Plane.BitDepth)
        Call Clamp(HL_ERR_Plane, sPlane1, "OPB_V1")

        Dim sPlane2 As CImgPlane
        Call GetFreePlane(sPlane2, sPlane1.planeGroup, sPlane1.BitDepth)
        Call MedianEx(sPlane1, sPlane2, "ZONE2D", 1, 5)

        Call MedianEx(sPlane1, sPlane2, "OPB_V1", 1, 5)

        Dim sPlane3 As CImgPlane
        Call GetFreePlane(sPlane3, sPlane2.planeGroup, sPlane2.BitDepth)
        Call MedianEx(sPlane2, sPlane3, "ZONE2D", 5, 1)

        Call MedianEx(sPlane2, sPlane3, "OPB_V1", 5, 1)

        '平均値取得
        Dim tmp1 As CImgColorAllResult
        Call AverageColorAll(sPlane3, "ZONE0", tmp1)

        Dim tmp2(nSite) As Double
        Call GetAverage_Color(tmp2, tmp1, "Gr1", "Gb1", "Gr2", "Gb2")

        Dim tmp3(nSite) As Double
        For site = 0 To nSite
                If TheExec.sites.site(site).Active Then
                        tmp3(site) = tmp2(site) * HL_ERR_LSB(site)
                End If
        Next site

        Call ResultAdd("HL_SEN", tmp3)

' #### HL_SEN_Z22 ####

        Dim tmp4 As CImgColorAllResult
        Call AverageColorAll(sPlane3, "ZONE22", tmp4)

        Dim tmp5(nSite) As Double
        Call GetAverage_Color(tmp5, tmp4, "Gr1", "Gb1", "Gr2", "Gb2")

        Dim tmp6(nSite) As Double
        For site = 0 To nSite
                If TheExec.sites.site(site).Active Then
                        tmp6(site) = tmp5(site) * HL_ERR_LSB(site)
                End If
        Next site

        Call ResultAdd("HL_SEN_Z22", tmp6)

' #### HL_SENGB1 ####

        Dim tmp7(nSite) As Double
        Call GetAverage_Color(tmp7, tmp1, "Gb1")

        Dim tmp8(nSite) As Double
        For site = 0 To nSite
                If TheExec.sites.site(site).Active Then
                        tmp8(site) = tmp7(site) * HL_ERR_LSB(site)
                End If
        Next site

        Call ResultAdd("HL_SENGB1", tmp8)

' #### HL_SENGB2 ####

        Dim tmp9(nSite) As Double
        Call GetAverage_Color(tmp9, tmp1, "Gb2")

        Dim tmp10(nSite) As Double
        For site = 0 To nSite
                If TheExec.sites.site(site).Active Then
                        tmp10(site) = tmp9(site) * HL_ERR_LSB(site)
                End If
        Next site

        Call ResultAdd("HL_SENGB2", tmp10)

' #### HL_SENGR1 ####

        Dim tmp11(nSite) As Double
        Call GetAverage_Color(tmp11, tmp1, "Gr1")

        Dim tmp12(nSite) As Double
        For site = 0 To nSite
                If TheExec.sites.site(site).Active Then
                        tmp12(site) = tmp11(site) * HL_ERR_LSB(site)
                End If
        Next site

        Call ResultAdd("HL_SENGR1", tmp12)

' #### HL_SENGR2 ####

        Dim tmp13(nSite) As Double
        Call GetAverage_Color(tmp13, tmp1, "Gr2")

        Dim tmp14(nSite) As Double
        For site = 0 To nSite
                If TheExec.sites.site(site).Active Then
                        tmp14(site) = tmp13(site) * HL_ERR_LSB(site)
                End If
        Next site

        Call ResultAdd("HL_SENGR2", tmp14)

' #### HL_SENR1 ####

        Dim tmp15(nSite) As Double
        Call GetAverage_Color(tmp15, tmp1, "R1")

        Dim tmp16(nSite) As Double
        For site = 0 To nSite
                If TheExec.sites.site(site).Active Then
                        tmp16(site) = tmp15(site) * HL_ERR_LSB(site)
                End If
        Next site

        Call ResultAdd("HL_SENR1", tmp16)

' #### HL_SENR2 ####

        Dim tmp17(nSite) As Double
        Call GetAverage_Color(tmp17, tmp1, "R2")

        Dim tmp18(nSite) As Double
        For site = 0 To nSite
                If TheExec.sites.site(site).Active Then
                        tmp18(site) = tmp17(site) * HL_ERR_LSB(site)
                End If
        Next site

        Call ResultAdd("HL_SENR2", tmp18)

' #### HL_SENB1 ####

        Dim tmp19(nSite) As Double
        Call GetAverage_Color(tmp19, tmp1, "B1")

        Dim tmp20(nSite) As Double
        For site = 0 To nSite
                If TheExec.sites.site(site).Active Then
                        tmp20(site) = tmp19(site) * HL_ERR_LSB(site)
                End If
        Next site

        Call ResultAdd("HL_SENB1", tmp20)

' #### HL_SENB2 ####

        Dim tmp21(nSite) As Double
        Call GetAverage_Color(tmp21, tmp1, "B2")

        Dim tmp22(nSite) As Double
        For site = 0 To nSite
                If TheExec.sites.site(site).Active Then
                        tmp22(site) = tmp21(site) * HL_ERR_LSB(site)
                End If
        Next site

        Call ResultAdd("HL_SENB2", tmp22)
        
'=== ↑貼り付けはここまで ===

End Sub


Attribute VB_Name = "Image_001_OF_ERR_Mod"

Option Explicit

Public Function OF_ERR_Process()

        Call PutImageInto_Common

' #### OF_QSMN_Z1 ####

    Dim site As Long

    ' 0.画像情報インポート.OF_QSMN_Z1
    Dim OF_ERR_Param As CParamPlane
    Dim OF_ERR_DevInfo As CDeviceConfigInfo
    Dim OF_ERR_Plane As CImgPlane
    Set OF_ERR_Param = TheParameterBank.Item("OFImageTest_Acq1")
    Set OF_ERR_DevInfo = TheDeviceProfiler.ConfigInfo("OFImageTest_Acq1")
        Call TheParameterBank.Delete("OFImageTest_Acq1")
    Set OF_ERR_Plane = OF_ERR_Param.plane

    ' 1.Clamp.OF_QSMN_Z1
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call Clamp(OF_ERR_Plane, sPlane1, "Bayer2x4_VOPB")

    ' 2.Median.OF_QSMN_Z1
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 1, 5)
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_VOPB", 1, 5)

    ' 3.Median.OF_QSMN_Z1
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call MedianEx(sPlane2, sPlane3, "Bayer2x4_ZONE3", 5, 1)
    Call MedianEx(sPlane2, sPlane3, "Bayer2x4_VOPB", 5, 1)

    ' 45.Min_FA.OF_QSMN_Z1
    Dim tmp1_0 As CImgColorAllResult
    Call Min_FA(sPlane3, "Bayer2x4_ZONE1", EEE_COLOR_ALL, tmp1_0)
        Call ReleasePlane(sPlane3)
    Dim tmp2 As CImgColorAllResult
    Call GetMin_CImgColor(tmp2, tmp1_0)

    ' 46.GetMin_Color.OF_QSMN_Z1
    Dim tmp3(nSite) As Double
    Call GetMin_Color(tmp3, tmp2, "-")

    ' 233.LSB定義.OF_QSMN_Z1
    Dim OF_ERR_LSB() As Double
     OF_ERR_LSB = OF_ERR_DevInfo.Lsb.AsDouble

    ' 240.LSB換算.OF_QSMN_Z1
    Dim tmp4(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp4(site) = tmp3(site) * OF_ERR_LSB(site)
        End If
    Next site

    ' 241.PutTestResult.OF_QSMN_Z1
    Call ResultAdd("OF_QSMN_Z1", tmp4)

' #### OF_4HLN ####

    ' 0.画像情報インポート.OF_4HLN

    ' 1.Clamp.OF_4HLN

    ' 2.Median.OF_4HLN
    Dim sPlane4 As CImgPlane
    Call GetFreePlane(sPlane4, "Normal_Bayer2x4", idpDepthS16, , "sPlane4")
    Call MedianEx(sPlane1, sPlane4, "Bayer2x4_ZONE3", 5, 1)
    Call MedianEx(sPlane1, sPlane4, "Bayer2x4_VOPB", 5, 1)

    ' 3.ZONE取得.OF_4HLN

    ' 5.AccumulateRow.OF_4HLN
    Dim sPlane5 As CImgPlane
    Call GetFreePlane(sPlane5, "Normal_Bayer2x4_ACR", idpDepthF32, , "sPlane5")
    Call MakeAcrPMD(sPlane5, "Bayer2x4_ZONE2D", "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT)
    Call AccumulateRow(sPlane4, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane5, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, idpAccumMean)
        Call ReleasePlane(sPlane4)

    ' 6.SubRows.OF_4HLN
    Dim sPlane6 As CImgPlane
    Call GetFreePlane(sPlane6, "Normal_Bayer2x4_ACR", idpDepthF32, , "sPlane6")
    Call SubRows(sPlane5, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, sPlane6, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, 4)
        Call ReleasePlane(sPlane5)
    Call MakeAcrJudgePMD(sPlane6, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", "Bayer2x4_ACR_4_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, 4)

    ' 9.AbsMax_FA.OF_4HLN
    Dim tmp5_0 As CImgColorAllResult
    Call AbsMax_FA(sPlane6, "Bayer2x4_ACR_4_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_ALL, tmp5_0)
    Dim tmp6 As CImgColorAllResult
    Call GetMax_CImgColor(tmp6, tmp5_0)

    ' 10.GetAbsMax_Color.OF_4HLN
    Dim tmp7(nSite) As Double
    Call GetAbsMax_Color(tmp7, tmp6, "-")

    ' 13.GetAbs.OF_4HLN
    Dim tmp8(nSite) As Double
    Call GetAbs(tmp8, tmp7)

    ' 14.パラメータ取得.OF_4HLN
    Dim tmp9(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp9(site) = 1
        End If
    Next site

    ' 15.計算式評価.OF_4HLN
    Dim tmp10(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp10(site) = Div(tmp8(site), tmp9(site), 999)
        End If
    Next site

    ' 16.d_read_vmcu_Line.OF_4HLN
    If (Flg_Debug = 1 Or Sw_Ana = 1) Then
        Call d_read_vmcu_Line(sPlane6, "Bayer2x4_ACR_4_ZONE2D_EEE_COLOR_FLAT", 1, OF_ERR_LSB, "NoKCO", "OF_HLINE", "mV", tmp7, "HLINE", "ABSMAX", tmp9)
    End If
        Call ReleasePlane(sPlane6)

    ' 17.LSB定義.OF_4HLN

    ' 18.LSB換算.OF_4HLN
    Dim tmp11(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp11(site) = tmp10(site) * OF_ERR_LSB(site)
        End If
    Next site

    ' 19.PutTestResult.OF_4HLN
    Call ResultAdd("OF_4HLN", tmp11)

' #### OF_VLN ####

    ' 0.画像情報インポート.OF_VLN

    ' 1.Clamp.OF_VLN

    ' 2.Median.OF_VLN

    ' 3.ZONE取得.OF_VLN

    ' 5.AccumulateColumn.OF_VLN
    Dim sPlane7 As CImgPlane
    Call GetFreePlane(sPlane7, "Normal_Bayer2x4_ACC", idpDepthF32, , "sPlane7")
    Call MakeAccPMD(sPlane7, "Bayer2x4_ZONE2D", "Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT)
    Call AccumulateColumn(sPlane2, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane7, "Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, idpAccumMean)
        Call ReleasePlane(sPlane2)

    ' 6.SubColumns.OF_VLN
    Dim sPlane8 As CImgPlane
    Call GetFreePlane(sPlane8, "Normal_Bayer2x4_ACC", idpDepthF32, , "sPlane8")
    Call SubColumns(sPlane7, "Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, sPlane8, "Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, 2)
        Call ReleasePlane(sPlane7)
    Call MakeAccJudgePMD(sPlane8, "Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", "Bayer2x4_ACC_2_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, 2)

    ' 7.AbsMax_FA.OF_VLN
    Dim tmp12_0 As CImgColorAllResult
    Call AbsMax_FA(sPlane8, "Bayer2x4_ACC_2_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_ALL, tmp12_0)
    Dim tmp13 As CImgColorAllResult
    Call GetMax_CImgColor(tmp13, tmp12_0)

    ' 8.GetAbsMax_Color.OF_VLN
    Dim tmp14(nSite) As Double
    Call GetAbsMax_Color(tmp14, tmp13, "-")

    ' 11.GetAbs.OF_VLN
    Dim tmp15(nSite) As Double
    Call GetAbs(tmp15, tmp14)

    ' 14.パラメータ取得.OF_VLN

    ' 15.計算式評価.OF_VLN
    Dim tmp16(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp16(site) = Div(tmp15(site), tmp9(site), 999)
        End If
    Next site

    ' 16.d_read_vmcu_Line.OF_VLN
    If (Flg_Debug = 1 Or Sw_Ana = 1) Then
        Call d_read_vmcu_Line(sPlane8, "Bayer2x4_ACC_2_ZONE2D_EEE_COLOR_FLAT", 1, OF_ERR_LSB, "NoKCO", "OF_VLINE", "mV", tmp14, "VLINE", "ABSMAX", tmp9)
    End If
        Call ReleasePlane(sPlane8)

    ' 17.LSB定義.OF_VLN

    ' 18.LSB換算.OF_VLN
    Dim tmp17(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp17(site) = tmp16(site) * OF_ERR_LSB(site)
        End If
    Next site

    ' 19.PutTestResult.OF_VLN
    Call ResultAdd("OF_VLN", tmp17)

' #### OF_FDL_Z2D ####

    ' 0.画像情報インポート.OF_FDL_Z2D

    ' 1.Clamp.OF_FDL_Z2D

    ' 2.LSB定義.OF_FDL_Z2D

    ' 3.SliceLevel生成.OF_FDL_Z2D
    Dim tmp_Slice1(nSite) As Double
    Call MakeSliceLevel(tmp_Slice1, 0.245, OF_ERR_LSB, , , , idpCountBelow)

'OF_FDL_Z2D OTPBLOW START -------------------------------------
'
'    ' 5.Count_FA.OF_FDL_Z2D
'    Dim tmp18_0 As CImgColorAllResult
'    Call count_FA(sPlane1, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountBelow, tmp_Slice1, tmp_Slice1, idpLimitEachSite, idpLimitExclude, tmp18_0, "FLG_OF_FDL_Z2D")
'    Dim tmp19 As CImgColorAllResult
'    Call GetSum_CImgColor(tmp19, tmp18_0)
'
'    ' 6.GetSum_Color.OF_FDL_Z2D
'    Dim tmp20(nSite) As Double
'    Call GetSum_Color(tmp20, tmp19, "-")
'
'    ' 7.PutTestResult.OF_FDL_Z2D
'    Call ResultAdd("OF_FDL_Z2D", tmp20)

    'Count_FA.OF_FDL_Z2D
    Dim tmp18_0 As CImgColorAllResult
    Call count_FA(sPlane1, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountBelow, tmp_Slice1, tmp_Slice1, idpLimitEachSite, idpLimitExclude, tmp18_0, "Flg_Temp1")

    'FlagCopy.OF_FDL_Z2D
    Dim sPlane12 As CImgPlane
    Call GetFreePlane(sPlane12, "Normal_Bayer2x4", idpDepthS16, True, "sPlane12")
    Call FlagCopy(sPlane12, "Bayer2x4_ZONE2D", "Flg_Temp1", 1)
        Call ClearALLFlagBit("Flg_Temp1")

    'Multimean.OF_FDL_Z2D
    Dim sPlane13 As CImgPlane
    Call GetFreePlane(sPlane13, "Normal_Bayer2x4_MUL", idpDepthS16, , "sPlane13")
    Call MakeMulPMD(sPlane13, "Bayer2x4_ZONE2D", "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", 2, 4, EEE_COLOR_FLAT)
    Call MultiMean(sPlane12, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane13, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", EEE_COLOR_FLAT, idpMultiMeanFuncSum, 2, 4)
        Call ReleasePlane(sPlane12)

    ' SliceLevel生成.OF_FDL_Z2D
    Dim tmp_Slice3(nSite) As Double
    Call MakeSliceLevel(tmp_Slice3, 1)

    'Count_FA.OF_FDL_Z2D
    Dim tmp30_0 As CImgColorAllResult
    Call count_FA(sPlane13, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", EEE_COLOR_ALL, idpCountAbove, tmp_Slice3, tmp_Slice3, idpLimitEachSite, idpLimitExclude, tmp30_0, "Flg_Temp2")
    Dim tmp31 As CImgColorAllResult
    Call GetSum_CImgColor(tmp31, tmp30_0)

    'GetSum_Color.OF_FDL_Z2D
    Dim tmp32(nSite) As Double
    Call GetSum_Color(tmp32, tmp31, "-")

    'readPixelSite.OF_FDL_Z2D
    Dim tmp_RPD2_0(nSite) As CPixInfo
    Call ReadPixelSite(sPlane13, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", tmp32, "Flg_Temp2", tmp_RPD2_0, idpAddrAbsolute)
        Call ClearALLFlagBit("Flg_Temp2")
        Call ReleasePlane(sPlane13)
    Dim tmp_RPD3(nSite) As CPixInfo
    Call RPDUnion(tmp_RPD3, tmp_RPD2_0)

    'RPDOffset.OF_FDL_Z2D
    Dim tmp_RPD4(nSite) As CPixInfo
    Call RPDOffset(tmp_RPD4, tmp_RPD3, -(2 - 1) + (1 - 1), -(4 - 1) + (3 - 1), 2, 4, 1)

    'MakeOtp.OF_FDL_Z2D
    Dim tmp2_Info_Hadd_OF_FDL_Z2D() As Double
    Dim tmp2_Info_Vadd_OF_FDL_Z2D() As Double
    Dim tmp2_Info_Dire_OF_FDL_Z2D() As Double
    Dim tmp2_Info_Sorc_OF_FDL_Z2D() As Double
    Dim tmp2_Info_Count_OF_FDL_Z2D(nSite) As Double
    Dim tmp33 As Double
    Dim i As Long
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If tmp33 < tmp_RPD4(site).Count Then
                tmp33 = tmp_RPD4(site).Count
            End If
        End If
    Next site
    ReDim tmp2_Info_Hadd_OF_FDL_Z2D(nSite, tmp33) As Double
    ReDim tmp2_Info_Vadd_OF_FDL_Z2D(nSite, tmp33) As Double
    ReDim tmp2_Info_Dire_OF_FDL_Z2D(nSite, tmp33) As Double
    ReDim tmp2_Info_Sorc_OF_FDL_Z2D(nSite, tmp33) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            For i = 0 To (tmp_RPD4(site).Count - 1)
                tmp2_Info_Hadd_OF_FDL_Z2D(site, i) = tmp_RPD4(site).PixInfo(i).x * 1 + (0)
                tmp2_Info_Vadd_OF_FDL_Z2D(site, i) = tmp_RPD4(site).PixInfo(i).y * 1 + (0)
                tmp2_Info_Dire_OF_FDL_Z2D(site, i) = 0
                tmp2_Info_Sorc_OF_FDL_Z2D(site, i) = 3
                If i > 1000 Then Exit For
            Next i
            tmp2_Info_Count_OF_FDL_Z2D(site) = tmp_RPD4(site).Count
        End If
    Next site

    'PutDefectResult.OF_FDL_Z2D
    Call ResultAdd("OF_FDL_Z2D_Info_Num", tmp2_Info_Count_OF_FDL_Z2D)
    Call ResultAdd("OF_FDL_Z2D_Info_Hadd", tmp2_Info_Hadd_OF_FDL_Z2D)
    Call ResultAdd("OF_FDL_Z2D_Info_Vadd", tmp2_Info_Vadd_OF_FDL_Z2D)
    Call ResultAdd("OF_FDL_Z2D_Info_Dire", tmp2_Info_Dire_OF_FDL_Z2D)
    Call ResultAdd("OF_FDL_Z2D_Info_Sorc", tmp2_Info_Sorc_OF_FDL_Z2D)

    'FD点欠陥フラグ取得(仮).OF_FDL_Z2D
    Dim tmp_RPD5(nSite) As CPixInfo
    Dim tmp_RPD6(nSite) As CPixInfo
    Dim AdrX As Double
    Dim AdrY As Double
    For AdrY = 0 To 4 - 1
        For AdrX = 0 To 2 - 1
            Call RPDOffset(tmp_RPD5, tmp_RPD4, AdrX, AdrY)
            Call RPDUnion(tmp_RPD6, tmp_RPD6, tmp_RPD5)
        Next AdrX
    Next AdrY

    'WritePixelAddrSite(PutFlag).OF_FDL_Z2D
    Dim sPlane14 As CImgPlane
    Call GetFreePlane(sPlane14, "Normal_Bayer2x4", idpDepthS16, True, "sPlane14")
    Call WritePixelAddrSite(sPlane14, "Bayer2x4_FULL", tmp_RPD6)

    'PutFlag_FA.OF_FDL_Z2D
    Call PutFlag_FA(sPlane14, "Bayer2x4_FULL", EEE_COLOR_ALL, idpCountAbove, 1, 1, idpLimitEachSite, idpLimitInclude, "FLG_OF_FDL_Z2D")
        Call ReleasePlane(sPlane14)

    'PutTestResult.OF_FDL_Z2D
    Call ResultAdd("OF_FDL_Z2D", tmp32)

'OF_FDL_Z2D OTPBLOW END -------------------------------------

' #### OF_ZL1 ####

    ' 0.画像情報インポート.OF_ZL1

    ' 1.Clamp.OF_ZL1

    ' 2.LSB定義.OF_ZL1

    ' 3.SliceLevel生成.OF_ZL1

    ' 4.マスク取得.OF_ZL1

    Call SharedFlagNot("Normal_Bayer2x4", "Bayer2x4_ZONE2D", "FLG_OF_FDL_Z2Dnot", "FLG_OF_FDL_Z2D")

    ' 5.Count_FA.OF_ZL1
    Dim tmp21_0 As CImgColorAllResult
    Call count_FA(sPlane1, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountBelow, tmp_Slice1, tmp_Slice1, idpLimitEachSite, idpLimitExclude, tmp21_0, "FLG_OF_ZL1", "FLG_OF_FDL_Z2Dnot")
    Dim tmp22 As CImgColorAllResult
    Call GetSum_CImgColor(tmp22, tmp21_0)

    ' 6.GetSum_Color.OF_ZL1
    Dim tmp23(nSite) As Double
    Call GetSum_Color(tmp23, tmp22, "-")

    ' 7.PutTestResult.OF_ZL1
    Call ResultAdd("OF_ZL1", tmp23)

'OF_ZL1 OTPBLOW START -------------------------------------
    
    'readPixelSite(OTP).OF_ZL1
    Dim tmp_RPD1_0(nSite) As CPixInfo
    Call ReadPixelSite(sPlane1, "Bayer2x4_ZONE2D", tmp23, "FLG_OF_ZL1", tmp_RPD1_0, idpAddrAbsolute)
    Dim tmp_RPD2(nSite) As CPixInfo
    Call RPDUnion(tmp_RPD2, tmp_RPD1_0)

    'MakeOtp.OF_ZL1
    Dim tmp1_Info_Hadd_OF_ZL1() As Double
    Dim tmp1_Info_Vadd_OF_ZL1() As Double
    Dim tmp1_Info_Dire_OF_ZL1() As Double
    Dim tmp1_Info_Sorc_OF_ZL1() As Double
    Dim tmp1_Info_Count_OF_ZL1(nSite) As Double
    Dim tmp30 As Double
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If tmp30 < tmp_RPD2(site).Count Then
                tmp30 = tmp_RPD2(site).Count
            End If
        End If
    Next site
    ReDim tmp1_Info_Hadd_OF_ZL1(nSite, tmp30) As Double
    ReDim tmp1_Info_Vadd_OF_ZL1(nSite, tmp30) As Double
    ReDim tmp1_Info_Dire_OF_ZL1(nSite, tmp30) As Double
    ReDim tmp1_Info_Sorc_OF_ZL1(nSite, tmp30) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            For i = 0 To (tmp_RPD2(site).Count - 1)
                tmp1_Info_Hadd_OF_ZL1(site, i) = tmp_RPD2(site).PixInfo(i).x * 1 + (0)
                tmp1_Info_Vadd_OF_ZL1(site, i) = tmp_RPD2(site).PixInfo(i).y * 1 + (0)
                tmp1_Info_Dire_OF_ZL1(site, i) = 0
                tmp1_Info_Sorc_OF_ZL1(site, i) = 1
                If i > 1000 Then Exit For
            Next i
            tmp1_Info_Count_OF_ZL1(site) = tmp_RPD2(site).Count
        End If
    Next site

    'PutDefectResult.OF_ZL1
    Call ResultAdd("OF_ZL1_Info_Num", tmp1_Info_Count_OF_ZL1)
    Call ResultAdd("OF_ZL1_Info_Hadd", tmp1_Info_Hadd_OF_ZL1)
    Call ResultAdd("OF_ZL1_Info_Vadd", tmp1_Info_Vadd_OF_ZL1)
    Call ResultAdd("OF_ZL1_Info_Dire", tmp1_Info_Dire_OF_ZL1)
    Call ResultAdd("OF_ZL1_Info_Sorc", tmp1_Info_Sorc_OF_ZL1)

'OF_ZL1 OTPBLOW END -------------------------------------


' #### DEFECT_1 ####

    ' 0.画像情報インポート.DEFECT_1

    ' 1.Clamp.DEFECT_1

    ' 2.LSB定義.DEFECT_1

    ' 3.SliceLevel生成.DEFECT_1

    ' 5.Count_FA.DEFECT_1
    Dim tmp24_0 As CImgColorAllResult
    Call count_FA(sPlane1, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountBelow, tmp_Slice1, tmp_Slice1, idpLimitEachSite, idpLimitExclude, tmp24_0)
    Dim tmp25 As CImgColorAllResult
    Call GetSum_CImgColor(tmp25, tmp24_0)

    ' 6.GetSum_Color.DEFECT_1
    Dim tmp26(nSite) As Double
    Call GetSum_Color(tmp26, tmp25, "-")

    ' 7.PutTestResult.DEFECT_1
    Call ResultAdd("DEFECT_1", tmp26)

    ' 8.d_read_vmcu_point.DEFECT_1
    If (Flg_Debug = 1 Or Sw_Ana = 1) Then
        Call d_read_vmcu_Point(sPlane1, "Bayer2x4_ZONE2D", 1000, OF_ERR_LSB, "NoKCO", "OF", "mV", idpCountBelow, tmp_Slice1, idpLimitExclude, "NoInputFlg", "-")
    End If
        Call ReleasePlane(sPlane1)

' #### DEFECT_2 ####

    ' 0.画像情報インポート.DEFECT_2

    ' 2.Median.DEFECT_2
    Dim sPlane9 As CImgPlane
    Call GetFreePlane(sPlane9, "Normal_Bayer2x4", idpDepthS16, , "sPlane9")
    Call MedianEx(OF_ERR_Plane, sPlane9, "Bayer2x4_ZONE3", 1, 5)
    Call MedianEx(OF_ERR_Plane, sPlane9, "Bayer2x4_VOPB", 1, 5)

    ' 3.Median.DEFECT_2
    Dim sPlane10 As CImgPlane
    Call GetFreePlane(sPlane10, "Normal_Bayer2x4", idpDepthS16, , "sPlane10")
    Call MedianEx(sPlane9, sPlane10, "Bayer2x4_ZONE3", 5, 1)
    Call MedianEx(sPlane9, sPlane10, "Bayer2x4_VOPB", 5, 1)
        Call ReleasePlane(sPlane9)

    ' 4.Subtract(通常).DEFECT_2
    Dim sPlane11 As CImgPlane
    Call GetFreePlane(sPlane11, "Normal_Bayer2x4", idpDepthS16, , "sPlane11")
    Call Subtract(OF_ERR_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane10, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane11, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane10)

    ' 5.LSB定義.DEFECT_2

    ' 6.SliceLevel生成.DEFECT_2
    Dim tmp_Slice2(nSite) As Double
    Call MakeSliceLevel(tmp_Slice2, 0.01, OF_ERR_LSB, , , 15 / 30, idpCountAbove)

    ' 8.Count_FA.DEFECT_2
    Dim tmp27_0 As CImgColorAllResult
    Call count_FA(sPlane11, "Bayer2x4_VOPB", EEE_COLOR_ALL, idpCountAbove, tmp_Slice2, tmp_Slice2, idpLimitEachSite, idpLimitExclude, tmp27_0)
    Dim tmp28 As CImgColorAllResult
    Call GetSum_CImgColor(tmp28, tmp27_0)

    ' 9.GetSum_Color.DEFECT_2
    Dim tmp29(nSite) As Double
    Call GetSum_Color(tmp29, tmp28, "-")

    ' 10.PutTestResult.DEFECT_2
    Call ResultAdd("DEFECT_2", tmp29)

    ' 14.d_read_vmcu_point.DEFECT_2
    If (Flg_Debug = 1 Or Sw_Ana = 1) Then
        Call d_read_vmcu_Point(sPlane11, "Bayer2x4_VOPB", 1000, OF_ERR_LSB, 15 / 30, "OPB", "mV", idpCountAbove, tmp_Slice2, idpLimitExclude, "NoInputFlg", "-")
    End If
        Call ReleasePlane(sPlane11)

End Function



Attribute VB_Name = "Image_003_DK18_ERR_Mod"

Option Explicit

Public Function DK18_ERR_Process()

        Call PutImageInto_Common

' #### DK_OBD_ZVOB ####

    Dim site As Long

    ' 0.画像情報インポート.DK_OBD_ZVOB
    Dim DK18_ERR_Param As CParamPlane
    Dim DK18_ERR_DevInfo As CDeviceConfigInfo
    Dim DK18_ERR_Plane As CImgPlane
    Set DK18_ERR_Param = TheParameterBank.Item("DK18ImageTest_Acq1")
    Set DK18_ERR_DevInfo = TheDeviceProfiler.ConfigInfo("DK18ImageTest_Acq1")
        Call TheParameterBank.Delete("DK18ImageTest_Acq1")
    Set DK18_ERR_Plane = DK18_ERR_Param.plane

    ' 2.Median.DK_OBD_ZVOB
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call MedianEx(DK18_ERR_Plane, sPlane1, "Bayer2x4_ZONE3", 1, 7)
    Call MedianEx(DK18_ERR_Plane, sPlane1, "Bayer2x4_VOPB", 1, 7)

    ' 3.Median.DK_OBD_ZVOB
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 7, 1)
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_VOPB", 7, 1)

    ' 228.Average_FA.DK_OBD_ZVOB
    Dim tmp1_0 As CImgColorAllResult
    Call Average_FA(sPlane2, "Bayer2x4_ZONEV1", EEE_COLOR_ALL, tmp1_0)
    Dim tmp2 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp2, tmp1_0)

    ' 229.GetAverage_Color.DK_OBD_ZVOB
    Dim tmp3(nSite) As Double
    Call GetAverage_Color(tmp3, tmp2, "-")

    ' 230.Average_FA.DK_OBD_ZVOB
    Dim tmp4_0 As CImgColorAllResult
    Call Average_FA(sPlane2, "Bayer2x4_VOPB", EEE_COLOR_ALL, tmp4_0)
    Dim tmp5 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp5, tmp4_0)

    ' 231.GetAverage_Color.DK_OBD_ZVOB
    Dim tmp6(nSite) As Double
    Call GetAverage_Color(tmp6, tmp5, "-")

    ' 232.計算式評価.DK_OBD_ZVOB
    Dim tmp7(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp7(site) = tmp3(site) - tmp6(site)
        End If
    Next site

    ' 233.LSB定義.DK_OBD_ZVOB
    Dim DK18_ERR_LSB() As Double
     DK18_ERR_LSB = DK18_ERR_DevInfo.Lsb.AsDouble

    ' 234.LSB換算.DK_OBD_ZVOB
    Dim tmp8(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp8(site) = tmp7(site) * 15 / 30 * DK18_ERR_LSB(site)
        End If
    Next site

    ' 235.PutTestResult.DK_OBD_ZVOB
    Call ResultAdd("DK_OBD_ZVOB", tmp8)

' #### DK_HLN ####

    ' 0.画像情報インポート.DK_HLN

    ' 2.Median.DK_HLN
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call MedianEx(DK18_ERR_Plane, sPlane3, "Bayer2x4_ZONE3", 7, 1)
    Call MedianEx(DK18_ERR_Plane, sPlane3, "Bayer2x4_VOPB", 7, 1)

    ' 3.ZONE取得.DK_HLN

    ' 5.AccumulateRow.DK_HLN
    Dim sPlane4 As CImgPlane
    Call GetFreePlane(sPlane4, "Normal_Bayer2x4_ACR", idpDepthF32, , "sPlane4")
    Call MakeAcrPMD(sPlane4, "Bayer2x4_ZONE2D", "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT)
    Call AccumulateRow(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane4, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, idpAccumMean)
        Call ReleasePlane(sPlane3)

    ' 6.SubRows.DK_HLN
    Dim sPlane5 As CImgPlane
    Call GetFreePlane(sPlane5, "Normal_Bayer2x4_ACR", idpDepthF32, , "sPlane5")
    Call SubRows(sPlane4, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, sPlane5, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, 2)
        Call ReleasePlane(sPlane4)
    Call MakeAcrJudgePMD(sPlane5, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", "Bayer2x4_ACR_2_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, 2)

    ' 9.AbsMax_FA.DK_HLN
    Dim tmp9_0 As CImgColorAllResult
    Call AbsMax_FA(sPlane5, "Bayer2x4_ACR_2_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_ALL, tmp9_0)
    Dim tmp10 As CImgColorAllResult
    Call GetMax_CImgColor(tmp10, tmp9_0)

    ' 10.GetAbsMax_Color.DK_HLN
    Dim tmp11(nSite) As Double
    Call GetAbsMax_Color(tmp11, tmp10, "-")

    ' 13.GetAbs.DK_HLN
    Dim tmp12(nSite) As Double
    Call GetAbs(tmp12, tmp11)

    ' 14.パラメータ取得.DK_HLN
    Dim tmp13(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp13(site) = 1
        End If
    Next site

    ' 15.計算式評価.DK_HLN
    Dim tmp14(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp14(site) = Div(tmp12(site), tmp13(site), 999)
        End If
    Next site

    ' 16.d_read_vmcu_Line.DK_HLN
    If (Flg_Debug = 1 Or Sw_Ana = 1) Then
        Call d_read_vmcu_Line(sPlane5, "Bayer2x4_ACR_2_ZONE2D_EEE_COLOR_FLAT", 1, DK18_ERR_LSB, 15 / 30, "DK_HLINE", "mV", tmp11, "HLINE", "ABSMAX", tmp13)
    End If
        Call ReleasePlane(sPlane5)

    ' 17.LSB定義.DK_HLN

    ' 18.LSB換算.DK_HLN
    Dim tmp15(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp15(site) = tmp14(site) * 15 / 30 * DK18_ERR_LSB(site)
        End If
    Next site

    ' 19.PutTestResult.DK_HLN
    Call ResultAdd("DK_HLN", tmp15)

' #### DK_VLN ####

    ' 0.画像情報インポート.DK_VLN

    ' 2.Median.DK_VLN

    ' 3.ZONE取得.DK_VLN

    ' 5.AccumulateColumn.DK_VLN
    Dim sPlane6 As CImgPlane
    Call GetFreePlane(sPlane6, "Normal_Bayer2x4_ACC", idpDepthF32, , "sPlane6")
    Call MakeAccPMD(sPlane6, "Bayer2x4_ZONE2D", "Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT)
    Call AccumulateColumn(sPlane1, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane6, "Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, idpAccumMean)
        Call ReleasePlane(sPlane1)

    ' 6.SubColumns.DK_VLN
    Dim sPlane7 As CImgPlane
    Call GetFreePlane(sPlane7, "Normal_Bayer2x4_ACC", idpDepthF32, , "sPlane7")
    Call SubColumns(sPlane6, "Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, sPlane7, "Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, 2)
        Call ReleasePlane(sPlane6)
    Call MakeAccJudgePMD(sPlane7, "Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", "Bayer2x4_ACC_2_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, 2)

    ' 7.AbsMax_FA.DK_VLN
    Dim tmp16_0 As CImgColorAllResult
    Call AbsMax_FA(sPlane7, "Bayer2x4_ACC_2_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_ALL, tmp16_0)
    Dim tmp17 As CImgColorAllResult
    Call GetMax_CImgColor(tmp17, tmp16_0)

    ' 8.GetAbsMax_Color.DK_VLN
    Dim tmp18(nSite) As Double
    Call GetAbsMax_Color(tmp18, tmp17, "-")

    ' 11.GetAbs.DK_VLN
    Dim tmp19(nSite) As Double
    Call GetAbs(tmp19, tmp18)

    ' 14.パラメータ取得.DK_VLN

    ' 15.計算式評価.DK_VLN
    Dim tmp20(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp20(site) = Div(tmp19(site), tmp13(site), 999)
        End If
    Next site

    ' 16.d_read_vmcu_Line.DK_VLN
    If (Flg_Debug = 1 Or Sw_Ana = 1) Then
        Call d_read_vmcu_Line(sPlane7, "Bayer2x4_ACC_2_ZONE2D_EEE_COLOR_FLAT", 1, DK18_ERR_LSB, 15 / 30, "DK_VLINE", "mV", tmp18, "VLINE", "ABSMAX", tmp13)
    End If
        Call ReleasePlane(sPlane7)

    ' 17.LSB定義.DK_VLN

    ' 18.LSB換算.DK_VLN
    Dim tmp21(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp21(site) = tmp20(site) * 15 / 30 * DK18_ERR_LSB(site)
        End If
    Next site

    ' 19.PutTestResult.DK_VLN
    Call ResultAdd("DK_VLN", tmp21)

' #### DK_ZV12_S1 ####

    ' 0.画像情報インポート.DK_ZV12_S1

    ' 2.Median.DK_ZV12_S1

    ' 3.Median.DK_ZV12_S1

    ' 4.Subtract(通常).DK_ZV12_S1
    Dim sPlane8 As CImgPlane
    Call GetFreePlane(sPlane8, "Normal_Bayer2x4", idpDepthS16, , "sPlane8")
    Call Subtract(DK18_ERR_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane2, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane8, "Bayer2x4_FULL", EEE_COLOR_ALL)

    ' 5.LSB定義.DK_ZV12_S1

    ' 6.SliceLevel生成.DK_ZV12_S1
    Dim tmp_Slice1(nSite) As Double
    Call MakeSliceLevel(tmp_Slice1, 0.001, DK18_ERR_LSB, , , 15 / 30, idpCountAbove)

    ' 8.Count_FA.DK_ZV12_S1
    Dim tmp22_0 As CImgColorAllResult
    Call count_FA(sPlane8, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice1, tmp_Slice1, idpLimitEachSite, idpLimitExclude, tmp22_0, "FLG_DK_ZV12_S1")
        Call ClearALLFlagBit("FLG_DK_ZV12_S1")
    Dim tmp23 As CImgColorAllResult
    Call GetSum_CImgColor(tmp23, tmp22_0)

    ' 9.GetSum_Color.DK_ZV12_S1
    Dim tmp24(nSite) As Double
    Call GetSum_Color(tmp24, tmp23, "-")

    ' 10.PutTestResult.DK_ZV12_S1
    Call ResultAdd("DK_ZV12_S1", tmp24)

' #### DK_ZL1IC ####

    ' 0.画像情報インポート.DK_ZL1IC

    ' 2.Median.DK_ZL1IC

    ' 3.Median.DK_ZL1IC

    ' 4.Subtract(通常).DK_ZL1IC

    ' 5.LSB定義.DK_ZL1IC

    ' 6.SliceLevel生成.DK_ZL1IC
    Dim tmp_Slice2(nSite) As Double
    Call MakeSliceLevel(tmp_Slice2, 0.0068, DK18_ERR_LSB, , , 15 / 30, idpCountAbove)

    ' 8.Count_FA.DK_ZL1IC
    Dim tmp25_0 As CImgColorAllResult
    Call count_FA(sPlane8, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice2, tmp_Slice2, idpLimitEachSite, idpLimitExclude, tmp25_0, "FLG_DK_ZL1IC")
    Dim tmp26 As CImgColorAllResult
    Call GetSum_CImgColor(tmp26, tmp25_0)

    ' 9.GetSum_Color.DK_ZL1IC
    Dim tmp27(nSite) As Double
    Call GetSum_Color(tmp27, tmp26, "-")

    ' 10.PutTestResult.DK_ZL1IC
    Call ResultAdd("DK_ZL1IC", tmp27)

    ' 14.d_read_vmcu_point.DK_ZL1IC
    If (Flg_Debug = 1 Or Sw_Ana = 1) Then
        Call d_read_vmcu_Point(sPlane8, "Bayer2x4_ZONE2D", 1000, DK18_ERR_LSB, 15 / 30, "DK_HOSEI_ZONE2D", "mV", idpCountAbove, tmp_Slice2, idpLimitExclude, "NoInputFlg", "-")
    End If

' #### HL_ZP6_S1 ####

    ' 0.画像情報インポート.HL_ZP6_S1
    Dim HL_ERR_Param As CParamPlane
    Dim HL_ERR_DevInfo As CDeviceConfigInfo
    Dim HL_ERR_Plane As CImgPlane
    Set HL_ERR_Param = TheParameterBank.Item("HLImageTest_Acq1")
    Set HL_ERR_DevInfo = TheDeviceProfiler.ConfigInfo("HLImageTest_Acq1")
        Call TheParameterBank.Delete("HLImageTest_Acq1")
    Set HL_ERR_Plane = HL_ERR_Param.plane

    ' 1.Clamp.HL_ZP6_S1
    Dim sPlane9 As CImgPlane
    Call GetFreePlane(sPlane9, "HL_TEMP_Bayer2x4", idpDepthS16, , "sPlane9")
    Call Clamp(HL_ERR_Plane, sPlane9, "HL_TEMP_Bayer2x4_VOPB")

    ' 2.Median.HL_ZP6_S1
    Dim sPlane10 As CImgPlane
    Call GetFreePlane(sPlane10, "HL_TEMP_Bayer2x4", idpDepthS16, , "sPlane10")
    Call MedianEx(sPlane9, sPlane10, "HL_TEMP_Bayer2x4_ZONE3", 1, 5)

    ' 3.Median.HL_ZP6_S1
    Dim sPlane11 As CImgPlane
    Call GetFreePlane(sPlane11, "HL_TEMP_Bayer2x4", idpDepthS16, , "sPlane11")
    Call MedianEx(sPlane10, sPlane11, "HL_TEMP_Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane10)

    ' 4.Copy.HL_ZP6_S1
    Dim sPlane12 As CImgPlane
    Call GetFreePlane(sPlane12, "HL_TEMP_Bayer2x4", idpDepthF32, , "sPlane12")
    Call Copy(sPlane11, "HL_TEMP_Bayer2x4_FULL", EEE_COLOR_ALL, sPlane12, "HL_TEMP_Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane11)

    ' 5.Copy.HL_ZP6_S1
    Dim sPlane13 As CImgPlane
    Call GetFreePlane(sPlane13, "HL_TEMP_Bayer2x4", idpDepthF32, , "sPlane13")
    Call Copy(sPlane9, "HL_TEMP_Bayer2x4_FULL", EEE_COLOR_ALL, sPlane13, "HL_TEMP_Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane9)

    ' 6.Divide(通常).HL_ZP6_S1
    Dim sPlane14 As CImgPlane
    Call GetFreePlane(sPlane14, "HL_TEMP_Bayer2x4", idpDepthF32, , "sPlane14")
    Call Divide(sPlane13, "HL_TEMP_Bayer2x4_FULL", EEE_COLOR_ALL, sPlane12, "HL_TEMP_Bayer2x4_FULL", EEE_COLOR_ALL, sPlane14, "HL_TEMP_Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane12)
        Call ReleasePlane(sPlane13)

    ' 7.LSB定義.HL_ZP6_S1
    Dim HL_ERR_LSB() As Double
     HL_ERR_LSB = HL_ERR_DevInfo.Lsb.AsDouble

    ' 8.SliceLevel生成.HL_ZP6_S1
    Dim tmp_Slice3(nSite) As Double
    Call MakeSliceLevel_Percent(tmp_Slice3, -0.06)

    ' 9.マスク取得.HL_ZP6_S1

    Call SharedFlagNot("Normal_Bayer2x4", "Bayer2x4_ZONE2D", "FLG_DK_ZL1ICnot", "FLG_DK_ZL1IC")

    ' 10.Count_FA.HL_ZP6_S1
    Dim tmp28_0 As CImgColorAllResult
    Call count_FA(sPlane14, "HL_TEMP_Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountBelow, tmp_Slice3, tmp_Slice3, idpLimitEachSite, idpLimitInclude, tmp28_0, "FLG_HL_ZP6_S1", "FLG_DK_ZL1ICnot")
        Call ClearALLFlagBit("FLG_HL_ZP6_S1")
    Dim tmp29 As CImgColorAllResult
    Call GetSum_CImgColor(tmp29, tmp28_0)

    ' 11.GetSum_Color.HL_ZP6_S1
    Dim tmp30(nSite) As Double
    Call GetSum_Color(tmp30, tmp29, "-")

    ' 12.PutTestResult.HL_ZP6_S1
    Call ResultAdd("HL_ZP6_S1", tmp30)

' #### HL_ZP6_S2 ####

    ' 0.画像情報インポート.HL_ZP6_S2

    ' 1.Clamp.HL_ZP6_S2

    ' 2.Median.HL_ZP6_S2

    ' 3.Median.HL_ZP6_S2

    ' 4.Copy.HL_ZP6_S2

    ' 5.Copy.HL_ZP6_S2

    ' 6.Divide(通常).HL_ZP6_S2

    ' 7.LSB定義.HL_ZP6_S2

    ' 8.SliceLevel生成.HL_ZP6_S2
    Dim tmp_Slice4(nSite) As Double
    Call MakeSliceLevel_Percent(tmp_Slice4, 0.06)

    ' 9.マスク取得.HL_ZP6_S2

    ' 10.Count_FA.HL_ZP6_S2
    Dim tmp31_0 As CImgColorAllResult
    Call count_FA(sPlane14, "HL_TEMP_Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice4, tmp_Slice4, idpLimitEachSite, idpLimitInclude, tmp31_0, "FLG_HL_ZP6_S2", "FLG_DK_ZL1ICnot")
        Call ClearALLFlagBit("FLG_HL_ZP6_S2")
    Dim tmp32 As CImgColorAllResult
    Call GetSum_CImgColor(tmp32, tmp31_0)

    ' 11.GetSum_Color.HL_ZP6_S2
    Dim tmp33(nSite) As Double
    Call GetSum_Color(tmp33, tmp32, "-")

    ' 12.PutTestResult.HL_ZP6_S2
    Call ResultAdd("HL_ZP6_S2", tmp33)

' #### HL_ZP6 ####

    ' 0.項目和.HL_ZP6
    Dim tmp_HL_ZP6_S1() As Double
    TheResult.GetResult "HL_ZP6_S1", tmp_HL_ZP6_S1
    Dim tmp_HL_ZP6_S2() As Double
    TheResult.GetResult "HL_ZP6_S2", tmp_HL_ZP6_S2
    Dim tmp34(nSite) As Double
    Call GetSum(tmp34, tmp_HL_ZP6_S1, tmp_HL_ZP6_S2)
    Dim tmp35(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            tmp35(site) = 1 + 1
        End If
    Next site

    ' 1.PutTestResult.HL_ZP6
    Call ResultAdd("HL_ZP6", tmp34)

' #### HL_ZP18_S1 ####

    ' 0.画像情報インポート.HL_ZP18_S1

    ' 1.Clamp.HL_ZP18_S1

    ' 2.Median.HL_ZP18_S1

    ' 3.Median.HL_ZP18_S1

    ' 4.Copy.HL_ZP18_S1

    ' 5.Copy.HL_ZP18_S1

    ' 6.Divide(通常).HL_ZP18_S1

    ' 7.LSB定義.HL_ZP18_S1

    ' 8.SliceLevel生成.HL_ZP18_S1
    Dim tmp_Slice5(nSite) As Double
    Call MakeSliceLevel_Percent(tmp_Slice5, -0.18)

    ' 9.マスク取得.HL_ZP18_S1

    ' 10.Count_FA.HL_ZP18_S1
    Dim tmp36_0 As CImgColorAllResult
    Call count_FA(sPlane14, "HL_TEMP_Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountBelow, tmp_Slice5, tmp_Slice5, idpLimitEachSite, idpLimitInclude, tmp36_0, "FLG_HL_ZP18_S1", "FLG_DK_ZL1ICnot")
    Dim tmp37 As CImgColorAllResult
    Call GetSum_CImgColor(tmp37, tmp36_0)

    ' 11.GetSum_Color.HL_ZP18_S1
    Dim tmp38(nSite) As Double
    Call GetSum_Color(tmp38, tmp37, "-")

    ' 12.PutTestResult.HL_ZP18_S1
    Call ResultAdd("HL_ZP18_S1", tmp38)

' #### HL_ZP18_S2 ####

    ' 0.画像情報インポート.HL_ZP18_S2

    ' 1.Clamp.HL_ZP18_S2

    ' 2.Median.HL_ZP18_S2

    ' 3.Median.HL_ZP18_S2

    ' 4.Copy.HL_ZP18_S2

    ' 5.Copy.HL_ZP18_S2

    ' 6.Divide(通常).HL_ZP18_S2

    ' 7.LSB定義.HL_ZP18_S2

    ' 8.SliceLevel生成.HL_ZP18_S2
    Dim tmp_Slice6(nSite) As Double
    Call MakeSliceLevel_Percent(tmp_Slice6, 0.18)

    ' 9.マスク取得.HL_ZP18_S2

    ' 10.Count_FA.HL_ZP18_S2
    Dim tmp39_0 As CImgColorAllResult
    Call count_FA(sPlane14, "HL_TEMP_Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice6, tmp_Slice6, idpLimitEachSite, idpLimitInclude, tmp39_0, "FLG_HL_ZP18_S2", "FLG_DK_ZL1ICnot")
    Dim tmp40 As CImgColorAllResult
    Call GetSum_CImgColor(tmp40, tmp39_0)

    ' 11.GetSum_Color.HL_ZP18_S2
    Dim tmp41(nSite) As Double
    Call GetSum_Color(tmp41, tmp40, "-")

    ' 12.PutTestResult.HL_ZP18_S2
    Call ResultAdd("HL_ZP18_S2", tmp41)

' #### HL_ZP18 ####

    ' 0.項目和.HL_ZP18
    Dim tmp_HL_ZP18_S1() As Double
    TheResult.GetResult "HL_ZP18_S1", tmp_HL_ZP18_S1
    Dim tmp_HL_ZP18_S2() As Double
    TheResult.GetResult "HL_ZP18_S2", tmp_HL_ZP18_S2
    Dim tmp42(nSite) As Double
    Call GetSum(tmp42, tmp_HL_ZP18_S1, tmp_HL_ZP18_S2)

    ' 1.PutTestResult.HL_ZP18
    Call ResultAdd("HL_ZP18", tmp42)

' #### HL_BZL4 ####

    ' 0.画像情報インポート.HL_BZL4

    ' 1.Clamp.HL_BZL4

    ' 2.Median.HL_BZL4

    ' 3.Median.HL_BZL4

    ' 4.Copy.HL_BZL4

    ' 5.Copy.HL_BZL4

    ' 6.Divide(通常).HL_BZL4

    ' 7.LSB定義.HL_BZL4

    ' 8.SliceLevel生成.HL_BZL4
    Dim tmp_Slice7(nSite) As Double
    Call MakeSliceLevel_Percent(tmp_Slice7, -0.8)

    ' 9.マスク取得.HL_BZL4

    Call SharedFlagOr("Normal_Bayer2x4", "Bayer2x4_ZONE2D", "FLG_DK_ZL1ICorOF_FDL_Z2D", "FLG_DK_ZL1IC", "FLG_OF_FDL_Z2D")

    Call SharedFlagOr("Normal_Bayer2x4", "Bayer2x4_ZONE2D", "FLG_DK_ZL1ICorOF_FDL_Z2DorOF_ZL1", "FLG_DK_ZL1ICorOF_FDL_Z2D", "FLG_OF_ZL1")
        Call ClearALLFlagBit("FLG_DK_ZL1ICorOF_FDL_Z2D")

    Call SharedFlagNot("Normal_Bayer2x4", "Bayer2x4_ZONE2D", "FLG_DK_ZL1ICorOF_FDL_Z2DorOF_ZL1not", "FLG_DK_ZL1ICorOF_FDL_Z2DorOF_ZL1")
        Call ClearALLFlagBit("FLG_DK_ZL1ICorOF_FDL_Z2DorOF_ZL1")

    ' 10.Count_FA.HL_BZL4
    Dim tmp43_0 As CImgColorAllResult
    Call count_FA(sPlane14, "HL_TEMP_Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountBelow, tmp_Slice7, tmp_Slice7, idpLimitEachSite, idpLimitInclude, tmp43_0, , "FLG_DK_ZL1ICorOF_FDL_Z2DorOF_ZL1not")
    Dim tmp44 As CImgColorAllResult
    Call GetSum_CImgColor(tmp44, tmp43_0)

    ' 11.GetSum_Color.HL_BZL4
    Dim tmp45(nSite) As Double
    Call GetSum_Color(tmp45, tmp44, "-")

    ' 12.PutTestResult.HL_BZL4
    Call ResultAdd("HL_BZL4", tmp45)

' #### HL_WZL4 ####

    ' 0.画像情報インポート.HL_WZL4

    ' 1.Clamp.HL_WZL4

    ' 2.Median.HL_WZL4

    ' 3.Median.HL_WZL4

    ' 4.Copy.HL_WZL4

    ' 5.Copy.HL_WZL4

    ' 6.Divide(通常).HL_WZL4

    ' 7.LSB定義.HL_WZL4

    ' 8.SliceLevel生成.HL_WZL4
    Dim tmp_Slice8(nSite) As Double
    Call MakeSliceLevel_Percent(tmp_Slice8, 0.8)

    ' 9.マスク取得.HL_WZL4

    ' 10.Count_FA.HL_WZL4
    Dim tmp46_0 As CImgColorAllResult
    Call count_FA(sPlane14, "HL_TEMP_Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice8, tmp_Slice8, idpLimitEachSite, idpLimitInclude, tmp46_0, , "FLG_DK_ZL1ICorOF_FDL_Z2DorOF_ZL1not")
        Call ClearALLFlagBit("FLG_DK_ZL1ICorOF_FDL_Z2DorOF_ZL1not")
    Dim tmp47 As CImgColorAllResult
    Call GetSum_CImgColor(tmp47, tmp46_0)

    ' 11.GetSum_Color.HL_WZL4
    Dim tmp48(nSite) As Double
    Call GetSum_Color(tmp48, tmp47, "-")

    ' 12.PutTestResult.HL_WZL4
    Call ResultAdd("HL_WZL4", tmp48)

' #### DEFECT_5 ####

    ' 0.画像情報インポート.DEFECT_5

    ' 1.Clamp.DEFECT_5

    ' 2.Median.DEFECT_5

    ' 3.Median.DEFECT_5

    ' 4.Copy.DEFECT_5

    ' 5.Copy.DEFECT_5

    ' 6.Divide(通常).DEFECT_5

    ' 7.LSB定義.DEFECT_5

    ' 8.SliceLevel生成.DEFECT_5
    Dim tmp_Slice9(nSite) As Double
    Call MakeSliceLevel_Percent(tmp_Slice9, -0.29)

    ' 9.マスク取得.DEFECT_5

    ' 10.Count_FA.DEFECT_5
    Dim tmp49_0 As CImgColorAllResult
    Call count_FA(sPlane14, "HL_TEMP_Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountBelow, tmp_Slice9, tmp_Slice9, idpLimitEachSite, idpLimitInclude, tmp49_0, , "FLG_DK_ZL1ICnot")
    Dim tmp50 As CImgColorAllResult
    Call GetSum_CImgColor(tmp50, tmp49_0)

    ' 11.GetSum_Color.DEFECT_5
    Dim tmp51(nSite) As Double
    Call GetSum_Color(tmp51, tmp50, "-")

    ' 12.PutTestResult.DEFECT_5
    Call ResultAdd("DEFECT_5", tmp51)

    ' 13.d_read_vmcu_point.DEFECT_5
    If (Flg_Debug = 1 Or Sw_Ana = 1) Then
        Call d_read_vmcu_Point(sPlane14, "HL_TEMP_Bayer2x4_ZONE2D", 1000, HL_ERR_LSB, "NoKCO", "HL_WZL2", "%", idpCountBelow, tmp_Slice9, idpLimitInclude, "FLG_DK_ZL1ICnot", "-")
    End If

' #### DEFECT_6 ####

    ' 0.画像情報インポート.DEFECT_6

    ' 1.Clamp.DEFECT_6

    ' 2.Median.DEFECT_6

    ' 3.Median.DEFECT_6

    ' 4.Copy.DEFECT_6

    ' 5.Copy.DEFECT_6

    ' 6.Divide(通常).DEFECT_6

    ' 7.LSB定義.DEFECT_6

    ' 8.SliceLevel生成.DEFECT_6
    Dim tmp_Slice10(nSite) As Double
    Call MakeSliceLevel_Percent(tmp_Slice10, 0.29)

    ' 9.マスク取得.DEFECT_6

    ' 10.Count_FA.DEFECT_6
    Dim tmp52_0 As CImgColorAllResult
    Call count_FA(sPlane14, "HL_TEMP_Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice10, tmp_Slice10, idpLimitEachSite, idpLimitInclude, tmp52_0, , "FLG_DK_ZL1ICnot")
    Dim tmp53 As CImgColorAllResult
    Call GetSum_CImgColor(tmp53, tmp52_0)

    ' 11.GetSum_Color.DEFECT_6
    Dim tmp54(nSite) As Double
    Call GetSum_Color(tmp54, tmp53, "-")

    ' 12.PutTestResult.DEFECT_6
    Call ResultAdd("DEFECT_6", tmp54)

    ' 13.d_read_vmcu_point.DEFECT_6
    If (Flg_Debug = 1 Or Sw_Ana = 1) Then
        Call d_read_vmcu_Point(sPlane14, "HL_TEMP_Bayer2x4_ZONE2D", 1000, HL_ERR_LSB, "NoKCO", "HLWB", "%", idpCountAbove, tmp_Slice10, idpLimitInclude, "FLG_DK_ZL1ICnot", "-")
    End If
        Call ReleasePlane(sPlane14)

' #### DK_ZV12_S2 ####

    ' 0.画像情報インポート.DK_ZV12_S2

    ' 2.Median.DK_ZV12_S2

    ' 3.Median.DK_ZV12_S2

    ' 4.Subtract(通常).DK_ZV12_S2

    ' 5.LSB定義.DK_ZV12_S2

    ' 6.SliceLevel生成.DK_ZV12_S2
    Dim tmp_Slice11(nSite) As Double
    Call MakeSliceLevel(tmp_Slice11, 0.002, DK18_ERR_LSB, , , 15 / 30, idpCountAbove)

    ' 8.Count_FA.DK_ZV12_S2
    Dim tmp55_0 As CImgColorAllResult
    Call count_FA(sPlane8, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice11, tmp_Slice11, idpLimitEachSite, idpLimitExclude, tmp55_0)
    Dim tmp56 As CImgColorAllResult
    Call GetSum_CImgColor(tmp56, tmp55_0)

    ' 9.GetSum_Color.DK_ZV12_S2
    Dim tmp57(nSite) As Double
    Call GetSum_Color(tmp57, tmp56, "-")

    ' 10.PutTestResult.DK_ZV12_S2
    Call ResultAdd("DK_ZV12_S2", tmp57)

' #### DK_ZV12 ####

    ' 0.測定結果取得.DK_ZV12
    Dim tmp_DK_ZV12_S1() As Double
    TheResult.GetResult "DK_ZV12_S1", tmp_DK_ZV12_S1

    ' 1.測定結果取得.DK_ZV12
    Dim tmp_DK_ZV12_S2() As Double
    TheResult.GetResult "DK_ZV12_S2", tmp_DK_ZV12_S2

    ' 2.計算式評価.DK_ZV12
    Dim tmp58(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp58(site) = tmp_DK_ZV12_S1(site) - tmp_DK_ZV12_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_ZV12
    Call ResultAdd("DK_ZV12", tmp58)

' #### DKH_FDL_Z2D ####

    ' 0.画像情報インポート.DKH_FDL_Z2D

    ' 2.Median.DKH_FDL_Z2D

    ' 3.Median.DKH_FDL_Z2D

    ' 4.Subtract(通常).DKH_FDL_Z2D

    ' 5.LSB定義.DKH_FDL_Z2D

    ' 6.SliceLevel生成.DKH_FDL_Z2D
    Dim tmp_Slice12(nSite) As Double
    Call MakeSliceLevel(tmp_Slice12, 0.004, DK18_ERR_LSB, , , , idpCountAbove)

    ' 7.マスク取得.DKH_FDL_Z2D

    Call SharedFlagNot("Normal_Bayer2x4", "Bayer2x4_FULL", "FLG_OF_FDL_Z2Dnot", "FLG_OF_FDL_Z2D")

    ' 8.FD共有CopyMask.DKH_FDL_Z2D
    Dim sPlane15 As CImgPlane
    Call GetFreePlane(sPlane15, "Normal_Bayer2x4", idpDepthS16, True, "sPlane15")
    Call Copy(sPlane8, sPlane8.BasePMD.Name, EEE_COLOR_FLAT, sPlane15, sPlane15.BasePMD.Name, EEE_COLOR_FLAT, "FLG_OF_FDL_Z2Dnot")
        Call ClearALLFlagBit("FLG_OF_FDL_Z2Dnot")

    ' 9.Multimean.DKH_FDL_Z2D
    Dim sPlane16 As CImgPlane
    Call GetFreePlane(sPlane16, "Normal_Bayer2x4_MUL", idpDepthS16, , "sPlane16")
    Call MakeMulPMD(sPlane16, "Bayer2x4_ZONE2D", "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", 2, 4, EEE_COLOR_FLAT)
    Call MultiMean(sPlane15, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane16, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", EEE_COLOR_FLAT, idpMultiMeanFuncMin, 2, 4)

    ' 10.Multimean.DKH_FDL_Z2D
    Dim sPlane17 As CImgPlane
    Call GetFreePlane(sPlane17, "Normal_Bayer2x4_MUL", idpDepthS16, , "sPlane17")
    Call MakeMulPMD(sPlane17, "Bayer2x4_ZONE2D", "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", 2, 4, EEE_COLOR_FLAT)
    Call MultiMean(sPlane15, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane17, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", EEE_COLOR_FLAT, idpMultiMeanFuncMax, 2, 4)
        Call ReleasePlane(sPlane15)

    ' 11.計算式評価.DKH_FDL_Z2D

    ' 12.計算式評価.DKH_FDL_Z2D
    Dim tmp59(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp59(site) = tmp13(site) / 1000
        End If
    Next site

    ' 13.SliceLevel生成.DKH_FDL_Z2D
    Dim tmp_Slice13(nSite) As Double
    Call MakeSliceLevel(tmp_Slice13, tmp59, DK18_ERR_LSB, , , , idpCountAbove)

    ' 14.PutFlag_FA.DKH_FDL_Z2D
    Call PutFlag_FA(sPlane16, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", EEE_COLOR_ALL, idpCountAbove, tmp_Slice13, tmp_Slice13, idpLimitEachSite, idpLimitExclude, "Flg_Temp1")
        Call ReleasePlane(sPlane16)

    ' 16.Count_FA.DKH_FDL_Z2D
    Dim tmp60_0 As CImgColorAllResult
    Call count_FA(sPlane17, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", EEE_COLOR_ALL, idpCountAbove, tmp_Slice12, tmp_Slice12, idpLimitEachSite, idpLimitExclude, tmp60_0, "Flg_Temp2", "Flg_Temp1")
    Dim tmp61 As CImgColorAllResult
    Call GetSum_CImgColor(tmp61, tmp60_0)

    ' 17.GetSum_Color.DKH_FDL_Z2D
    Dim tmp62(nSite) As Double
    Call GetSum_Color(tmp62, tmp61, "-")

    ' 18.readPixelSite.DKH_FDL_Z2D
    Dim tmp_RPD1_0(nSite) As CPixInfo
    Call ReadPixelSite(sPlane17, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", tmp62, "Flg_Temp2", tmp_RPD1_0, idpAddrAbsolute)
        Call ReleasePlane(sPlane17)
    Dim tmp_RPD2(nSite) As CPixInfo
    Call RPDUnion(tmp_RPD2, tmp_RPD1_0)

    ' 19.RPDOffset.DKH_FDL_Z2D
    Dim tmp_RPD3(nSite) As CPixInfo
    Call RPDOffset(tmp_RPD3, tmp_RPD2, -(2 - 1) + (1 - 1), -(4 - 1) + (3 - 1), 2, 4, 1)

    ' 20.MakeOtp.DKH_FDL_Z2D
    Dim tmp13_Info_Hadd_DKH_FDL_Z2D() As Double
    Dim tmp13_Info_Vadd_DKH_FDL_Z2D() As Double
    Dim tmp13_Info_Dire_DKH_FDL_Z2D() As Double
    Dim tmp13_Info_Sorc_DKH_FDL_Z2D() As Double
    Dim tmp13_Info_Count_DKH_FDL_Z2D(nSite) As Double
    Dim tmp63 As Double
    Dim i As Long
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If tmp63 < tmp_RPD3(site).Count Then
                tmp63 = tmp_RPD3(site).Count
            End If
        End If
    Next site
    ReDim tmp13_Info_Hadd_DKH_FDL_Z2D(nSite, tmp63) As Double
    ReDim tmp13_Info_Vadd_DKH_FDL_Z2D(nSite, tmp63) As Double
    ReDim tmp13_Info_Dire_DKH_FDL_Z2D(nSite, tmp63) As Double
    ReDim tmp13_Info_Sorc_DKH_FDL_Z2D(nSite, tmp63) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            For i = 0 To (tmp_RPD3(site).Count - 1)
                tmp13_Info_Hadd_DKH_FDL_Z2D(site, i) = tmp_RPD3(site).PixInfo(i).x * 1 + (0)
                tmp13_Info_Vadd_DKH_FDL_Z2D(site, i) = tmp_RPD3(site).PixInfo(i).y * 1 + (0)
                tmp13_Info_Dire_DKH_FDL_Z2D(site, i) = 0
                tmp13_Info_Sorc_DKH_FDL_Z2D(site, i) = 3
            Next i
            tmp13_Info_Count_DKH_FDL_Z2D(site) = tmp_RPD3(site).Count
        End If
    Next site

    ' 21.PutDefectResult.DKH_FDL_Z2D
    Call ResultAdd("DKH_FDL_Z2D_Info_Num", tmp13_Info_Count_DKH_FDL_Z2D)
    Call ResultAdd("DKH_FDL_Z2D_Info_Hadd", tmp13_Info_Hadd_DKH_FDL_Z2D)
    Call ResultAdd("DKH_FDL_Z2D_Info_Vadd", tmp13_Info_Vadd_DKH_FDL_Z2D)
    Call ResultAdd("DKH_FDL_Z2D_Info_Dire", tmp13_Info_Dire_DKH_FDL_Z2D)
    Call ResultAdd("DKH_FDL_Z2D_Info_Sorc", tmp13_Info_Sorc_DKH_FDL_Z2D)

    ' 22.FD点欠陥フラグ取得(仮).DKH_FDL_Z2D
    Dim tmp_RPD4(nSite) As CPixInfo
    Dim tmp_RPD5(nSite) As CPixInfo
    Dim AdrX As Double
    Dim AdrY As Double
    For AdrY = 0 To 4 - 1
        For AdrX = 0 To 2 - 1
            Call RPDOffset(tmp_RPD4, tmp_RPD3, AdrX, AdrY)
            Call RPDUnion(tmp_RPD5, tmp_RPD5, tmp_RPD4)
        Next AdrX
    Next AdrY

    ' 23.WritePixelAddrSite(PutFlag).DKH_FDL_Z2D
    Dim sPlane18 As CImgPlane
    Call GetFreePlane(sPlane18, "Normal_Bayer2x4", idpDepthS16, True, "sPlane18")
    Call WritePixelAddrSite(sPlane18, "Bayer2x4_FULL", tmp_RPD5)

    ' 24.PutFlag_FA.DKH_FDL_Z2D
    Call PutFlag_FA(sPlane18, "Bayer2x4_FULL", EEE_COLOR_ALL, idpCountAbove, 1, 1, idpLimitEachSite, idpLimitInclude, "FLG_DKH_FDL_Z2D")
        Call ReleasePlane(sPlane18)

    ' 25.PutTestResult.DKH_FDL_Z2D
    Call ResultAdd("DKH_FDL_Z2D", tmp62)

' #### DK_SH ####

    ' 0.画像情報インポート.DK_SH

    ' 2.Median.DK_SH

    ' 3.Median.DK_SH

    ' 5.ZONE取得.DK_SH

    ' 7.Multimean.DK_SH
    Dim sPlane19 As CImgPlane
    Call GetFreePlane(sPlane19, "Normal_Bayer2x4_MUL", idpDepthF32, , "sPlane19")
    Call MakeMulPMD(sPlane19, "Bayer2x4_ZONE2D", "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_10_8", 10, 8, EEE_COLOR_FLAT)
    Call MultiMean(sPlane2, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane19, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_10_8", EEE_COLOR_FLAT, idpMultiMeanFuncMean, 10, 8)
        Call ReleasePlane(sPlane2)

    ' 8.Min_FA.DK_SH
    Dim tmp64_0 As CImgColorAllResult
    Call Min_FA(sPlane19, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_10_8", EEE_COLOR_ALL, tmp64_0)
    Dim tmp65 As CImgColorAllResult
    Call GetMin_CImgColor(tmp65, tmp64_0)

    ' 9.GetMin_Color.DK_SH
    Dim tmp66(nSite) As Double
    Call GetMin_Color(tmp66, tmp65, "-")

    ' 10.Max_FA.DK_SH
    Dim tmp67_0 As CImgColorAllResult
    Call Max_FA(sPlane19, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_10_8", EEE_COLOR_ALL, tmp67_0)
        Call ReleasePlane(sPlane19)
    Dim tmp68 As CImgColorAllResult
    Call GetMax_CImgColor(tmp68, tmp67_0)

    ' 11.GetMax_Color.DK_SH
    Dim tmp69(nSite) As Double
    Call GetMax_Color(tmp69, tmp68, "-")

    ' 12.計算式評価.DK_SH
    Dim tmp70(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp70(site) = tmp69(site) - tmp66(site)
        End If
    Next site

    ' 13.パラメータ取得.DK_SH

    ' 14.計算式評価.DK_SH
    Dim tmp71(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp71(site) = Div(tmp70(site), tmp13(site), 999)
        End If
    Next site

    ' 15.LSB定義.DK_SH

    ' 16.LSB換算.DK_SH
    Dim tmp72(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp72(site) = tmp71(site) * DK18_ERR_LSB(site)
        End If
    Next site

    ' 17.PutTestResult.DK_SH
    Call ResultAdd("DK_SH", tmp72)

' #### DK_SIG ####

    ' 0.画像情報インポート.DK_SIG

    ' 1.Clamp.DK_SIG
    Dim sPlane20 As CImgPlane
    Call GetFreePlane(sPlane20, "Normal_Bayer2x4", idpDepthS16, , "sPlane20")
    Call SubtractConst(DK18_ERR_Plane, DK18_ERR_Plane.BasePMD.Name, EEE_COLOR_ALL, 63, sPlane20, DK18_ERR_Plane.BasePMD.Name, EEE_COLOR_ALL)

    ' 2.Median.DK_SIG
    Dim sPlane21 As CImgPlane
    Call GetFreePlane(sPlane21, "Normal_Bayer2x4", idpDepthS16, , "sPlane21")
    Call MedianEx(sPlane20, sPlane21, "Bayer2x4_ZONE3", 1, 7)
    Call MedianEx(sPlane20, sPlane21, "Bayer2x4_VOPB", 1, 7)
        Call ReleasePlane(sPlane20)

    ' 3.Median.DK_SIG
    Dim sPlane22 As CImgPlane
    Call GetFreePlane(sPlane22, "Normal_Bayer2x4", idpDepthS16, , "sPlane22")
    Call MedianEx(sPlane21, sPlane22, "Bayer2x4_ZONE3", 7, 1)
    Call MedianEx(sPlane21, sPlane22, "Bayer2x4_VOPB", 7, 1)
        Call ReleasePlane(sPlane21)

    ' 82.Average_FA.DK_SIG
    Dim tmp73_0 As CImgColorAllResult
    Call Average_FA(sPlane22, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, tmp73_0)
    Dim tmp74 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp74, tmp73_0)

    ' 83.GetAverage_Color.DK_SIG
    Dim tmp75(nSite) As Double
    Call GetAverage_Color(tmp75, tmp74, "-")

    ' 233.LSB定義.DK_SIG

    ' 238.LSB換算.DK_SIG
    Dim tmp76(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp76(site) = tmp75(site) * 15 / 30 * DK18_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.DK_SIG
    Call ResultAdd("DK_SIG", tmp76)

' #### DK_SIG_ZVOB ####

    ' 0.画像情報インポート.DK_SIG_ZVOB

    ' 1.Clamp.DK_SIG_ZVOB

    ' 2.Median.DK_SIG_ZVOB

    ' 3.Median.DK_SIG_ZVOB

    ' 82.Average_FA.DK_SIG_ZVOB
    Dim tmp77_0 As CImgColorAllResult
    Call Average_FA(sPlane22, "Bayer2x4_VOPB", EEE_COLOR_ALL, tmp77_0)
        Call ReleasePlane(sPlane22)
    Dim tmp78 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp78, tmp77_0)

    ' 83.GetAverage_Color.DK_SIG_ZVOB
    Dim tmp79(nSite) As Double
    Call GetAverage_Color(tmp79, tmp78, "-")

    ' 233.LSB定義.DK_SIG_ZVOB

    ' 238.LSB換算.DK_SIG_ZVOB
    Dim tmp80(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp80(site) = tmp79(site) * 15 / 30 * DK18_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.DK_SIG_ZVOB
    Call ResultAdd("DK_SIG_ZVOB", tmp80)

' #### DK_SIGDF ####

    ' 0.測定結果取得.DK_SIGDF
    Dim tmp_DK_SIG() As Double
    TheResult.GetResult "DK_SIG", tmp_DK_SIG

    ' 1.測定結果取得.DK_SIGDF
    Dim tmp_DK_SIG_ZVOB() As Double
    TheResult.GetResult "DK_SIG_ZVOB", tmp_DK_SIG_ZVOB

    ' 2.計算式評価.DK_SIGDF
    Dim tmp81(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp81(site) = tmp_DK_SIG(site) - tmp_DK_SIG_ZVOB(site)
        End If
    Next site

    ' 3.PutTestResult.DK_SIGDF
    Call ResultAdd("DK_SIGDF", tmp81)

' #### DEFECT_3 ####

    ' 0.画像情報インポート.DEFECT_3

    ' 2.Median.DEFECT_3

    ' 3.Median.DEFECT_3

    ' 4.Subtract(通常).DEFECT_3

    ' 5.LSB定義.DEFECT_3

    ' 6.SliceLevel生成.DEFECT_3

    ' 8.Count_FA.DEFECT_3
    Dim tmp82_0 As CImgColorAllResult
    Call count_FA(sPlane8, "Bayer2x4_VOPB", EEE_COLOR_ALL, idpCountAbove, tmp_Slice2, tmp_Slice2, idpLimitEachSite, idpLimitExclude, tmp82_0)
    Dim tmp83 As CImgColorAllResult
    Call GetSum_CImgColor(tmp83, tmp82_0)

    ' 9.GetSum_Color.DEFECT_3
    Dim tmp84(nSite) As Double
    Call GetSum_Color(tmp84, tmp83, "-")

    ' 10.PutTestResult.DEFECT_3
    Call ResultAdd("DEFECT_3", tmp84)

    ' 14.d_read_vmcu_point.DEFECT_3
    If (Flg_Debug = 1 Or Sw_Ana = 1) Then
        Call d_read_vmcu_Point(sPlane8, "Bayer2x4_VOPB", 1000, DK18_ERR_LSB, 15 / 30, "DK_HOSEI_VOPB", "mV", idpCountAbove, tmp_Slice2, idpLimitExclude, "NoInputFlg", "-")
    End If

' #### DEFECT_4 ####

    ' 0.画像情報インポート.DEFECT_4

    ' 2.Median.DEFECT_4

    ' 3.Median.DEFECT_4

    ' 4.Subtract(通常).DEFECT_4

    ' 5.LSB定義.DEFECT_4

    ' 6.SliceLevel生成.DEFECT_4
    Dim tmp_Slice14(nSite) As Double
    Call MakeSliceLevel(tmp_Slice14, 0.0008, DK18_ERR_LSB, , , , idpCountAbove)

    ' 8.FD共有CopyMask.DEFECT_4
    Dim sPlane23 As CImgPlane
    Call GetFreePlane(sPlane23, "Normal_Bayer2x4", idpDepthS16, True, "sPlane23")
    Call Copy(sPlane8, sPlane8.BasePMD.Name, EEE_COLOR_FLAT, sPlane23, sPlane23.BasePMD.Name, EEE_COLOR_FLAT)
        Call ReleasePlane(sPlane8)

    ' 9.Multimean.DEFECT_4
    Dim sPlane24 As CImgPlane
    Call GetFreePlane(sPlane24, "Normal_Bayer2x4_MUL", idpDepthS16, , "sPlane24")
    Call MakeMulPMD(sPlane24, "Bayer2x4_ZONE2D", "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", 2, 4, EEE_COLOR_FLAT)
    Call MultiMean(sPlane23, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane24, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", EEE_COLOR_FLAT, idpMultiMeanFuncMin, 2, 4)

    ' 10.Multimean.DEFECT_4
    Dim sPlane25 As CImgPlane
    Call GetFreePlane(sPlane25, "Normal_Bayer2x4_MUL", idpDepthS16, , "sPlane25")
    Call MakeMulPMD(sPlane25, "Bayer2x4_ZONE2D", "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", 2, 4, EEE_COLOR_FLAT)
    Call MultiMean(sPlane23, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane25, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", EEE_COLOR_FLAT, idpMultiMeanFuncMax, 2, 4)

    ' 11.計算式評価.DEFECT_4
    Dim tmp85(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp85(site) = 0.8
        End If
    Next site

    ' 12.計算式評価.DEFECT_4
    Dim tmp86(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp86(site) = tmp85(site) / 1000
        End If
    Next site

    ' 13.SliceLevel生成.DEFECT_4
    Dim tmp_Slice15(nSite) As Double
    Call MakeSliceLevel(tmp_Slice15, tmp86, DK18_ERR_LSB, , , , idpCountAbove)

    ' 14.PutFlag_FA.DEFECT_4
    Call PutFlag_FA(sPlane24, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", EEE_COLOR_ALL, idpCountAbove, tmp_Slice15, tmp_Slice15, idpLimitEachSite, idpLimitExclude, "Flg_Temp1")
        Call ReleasePlane(sPlane24)

    ' 15.d_read_vmcu_FD.DEFECT_4
    If (Flg_Debug = 1 Or Sw_Ana = 1) Then
        Dim tmp87_0 As CImgColorAllResult
        Call count_FA(sPlane25, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", EEE_COLOR_ALL, idpCountAbove, tmp_Slice14, tmp_Slice14, idpLimitEachSite, idpLimitExclude, tmp87_0, "FLG_FD_Point_DEFECT_4", "Flg_Temp1")
        Dim tmp88 As CImgColorAllResult
        Call GetSum_CImgColor(tmp88, tmp87_0)
        Dim tmp89(nSite) As Double
        Call GetSum_Color(tmp89, tmp88, "-")
        Call d_read_vmcu_FD(sPlane23, sPlane25, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", 1000, DK18_ERR_LSB, "NoKCO", "DKH_FDL_Z2D_08MV", "mV", tmp89, "FLG_FD_Point_DEFECT_4", -(2 - 1) + (1 - 1), -(4 - 1) + (3 - 1), 2, 4)
    End If
        Call ClearALLFlagBit("FLG_FD_Point_DEFECT_4")
        Call ReleasePlane(sPlane23)

    ' 16.Count_FA.DEFECT_4
    Dim tmp90_0 As CImgColorAllResult
    Call count_FA(sPlane25, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", EEE_COLOR_ALL, idpCountAbove, tmp_Slice14, tmp_Slice14, idpLimitEachSite, idpLimitExclude, tmp90_0, "Flg_Temp2", "Flg_Temp1")
    Dim tmp91 As CImgColorAllResult
    Call GetSum_CImgColor(tmp91, tmp90_0)

    ' 17.GetSum_Color.DEFECT_4
    Dim tmp92(nSite) As Double
    Call GetSum_Color(tmp92, tmp91, "-")

    ' 18.readPixelSite.DEFECT_4
    Dim tmp_RPD6_0(nSite) As CPixInfo
    Call ReadPixelSite(sPlane25, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", tmp92, "Flg_Temp2", tmp_RPD6_0, idpAddrAbsolute)
        Call ReleasePlane(sPlane25)
    Dim tmp_RPD7(nSite) As CPixInfo
    Call RPDUnion(tmp_RPD7, tmp_RPD6_0)

    ' 19.RPDOffset.DEFECT_4
    Dim tmp_RPD8(nSite) As CPixInfo
    Call RPDOffset(tmp_RPD8, tmp_RPD7, -(2 - 1) + (1 - 1), -(4 - 1) + (3 - 1), 2, 4, 1)

    ' 25.PutTestResult.DEFECT_4
    Call ResultAdd("DEFECT_4", tmp92)

' #### HLD_33F3SC ####

    ' 0.フラグ取得.HLD_33F3SC
    Dim tmp_DK_ZL1IC() As Double
    TheResult.GetResult "DK_ZL1IC", tmp_DK_ZL1IC
    Dim tmp_DKH_FDL_Z2D() As Double
    TheResult.GetResult "DKH_FDL_Z2D", tmp_DKH_FDL_Z2D
    Dim tmp_OF_FDL_Z2D() As Double
    TheResult.GetResult "OF_FDL_Z2D", tmp_OF_FDL_Z2D
    Dim tmp_OF_ZL1() As Double
    TheResult.GetResult "OF_ZL1", tmp_OF_ZL1
    Call MakeOrPMD("Normal_Bayer2x2", "Bayer2x2_ZONE2D", "Bayer2x2_ZONE2D")
    Call SharedFlagOr_Array("Normal_Bayer2x2", "Bayer2x2_ZONE2D", "FLG_DK_ZL1IC_DKH_FDL_Z2D_HL_ZP18_S1_HL_ZP18_S2_OF_FDL_Z2D_OF_ZL1_Bayer2x2_ZONE2D", "FLG_DK_ZL1IC", "FLG_DKH_FDL_Z2D", "FLG_HL_ZP18_S1", "FLG_HL_ZP18_S2", "FLG_OF_FDL_Z2D", "FLG_OF_ZL1")

    ' 3.FlagCopy.HLD_33F3SC
    Dim sPlane26 As CImgPlane
    Call GetFreePlane(sPlane26, "Normal_Bayer2x2", idpDepthS16, True, "sPlane26")
    Call FlagCopy(sPlane26, "Bayer2x2_ZONE2D", "FLG_DK_ZL1IC_DKH_FDL_Z2D_HL_ZP18_S1_HL_ZP18_S2_OF_FDL_Z2D_OF_ZL1_Bayer2x2_ZONE2D", 1)
        Call ClearALLFlagBit("FLG_DK_ZL1IC_DKH_FDL_Z2D_HL_ZP18_S1_HL_ZP18_S2_OF_FDL_Z2D_OF_ZL1_Bayer2x2_ZONE2D")

    ' 5.Convolution.HLD_33F3SC
    Dim sPlane27 As CImgPlane
    Call GetFreePlane(sPlane27, "Normal_Bayer2x2", idpDepthS16, , "sPlane27")
    Call Convolution(sPlane26, "Bayer2x2_ZONE2D", EEE_COLOR_ALL, sPlane27, "Bayer2x2_ZONE2D", EEE_COLOR_ALL, "kernel_Cluster")
        Call ReleasePlane(sPlane26)

    ' 6.Count_FA.HLD_33F3SC
    Dim tmp93_0 As CImgColorAllResult
    Call count_FA(sPlane27, "Bayer2x2_ZONE2D", EEE_COLOR_ALL, idpCountAbove, 5, 5, idpLimitEachSite, idpLimitExclude, tmp93_0)
        Call ReleasePlane(sPlane27)
    Dim tmp94 As CImgColorAllResult
    Call GetSum_CImgColor(tmp94, tmp93_0)

    ' 7.GetSum_Color.HLD_33F3SC
    Dim tmp95(nSite) As Double
    Call GetSum_Color(tmp95, tmp94, "-")

    ' 8.PutTestResult.HLD_33F3SC
    Call ResultAdd("HLD_33F3SC", tmp95)

' #### HLD_33FSC ####

    ' 0.フラグ取得.HLD_33FSC
    Call SharedFlagOr_Array("Normal_Bayer2x2", "Bayer2x2_ZONE2D", "FLG_DK_ZL1IC_HL_ZP18_S1_HL_ZP18_S2_Bayer2x2_ZONE2D", "FLG_DK_ZL1IC", "FLG_HL_ZP18_S1", "FLG_HL_ZP18_S2")
        Call ClearALLFlagBit("FLG_HL_ZP18_S2")
        Call ClearALLFlagBit("FLG_HL_ZP18_S1")

    ' 1.GetSum.HLD_33FSC
    Dim tmp96(nSite) As Double
    Call GetSum(tmp96, tmp_DK_ZL1IC, tmp_HL_ZP18_S1, tmp_HL_ZP18_S2)

    ' 2.GetMaxArr.HLD_33FSC
    Dim tmp97 As Double
    tmp97 = GetMaxArr(tmp96)

    ' 3.マスク取得.HLD_33FSC

    Call SharedFlagOr("Normal_Bayer2x4", "Bayer2x4_FULL", "FLG_DKH_FDL_Z2DorOF_FDL_Z2D", "FLG_DKH_FDL_Z2D", "FLG_OF_FDL_Z2D")
        Call ClearALLFlagBit("FLG_OF_FDL_Z2D")
        Call ClearALLFlagBit("FLG_DKH_FDL_Z2D")

    Call SharedFlagOr("Normal_Bayer2x4", "Bayer2x4_FULL", "FLG_DKH_FDL_Z2DorOF_FDL_Z2DorOF_ZL1", "FLG_DKH_FDL_Z2DorOF_FDL_Z2D", "FLG_OF_ZL1")
        Call ClearALLFlagBit("FLG_OF_ZL1")
        Call ClearALLFlagBit("FLG_DKH_FDL_Z2DorOF_FDL_Z2D")

    Call SharedFlagNot("Normal_Bayer2x4", "Bayer2x4_FULL", "FLG_DKH_FDL_Z2DorOF_FDL_Z2DorOF_ZL1not", "FLG_DKH_FDL_Z2DorOF_FDL_Z2DorOF_ZL1")
        Call ClearALLFlagBit("FLG_DKH_FDL_Z2DorOF_FDL_Z2DorOF_ZL1")

    ' 4.SharedFlagAnd.HLD_33FSC
Call ClearALLFlagBit("Flg_Temp1")
    Call SharedFlagAnd("Normal_Bayer2x2", "Bayer2x2_FULL", "Flg_Temp1", "FLG_DK_ZL1IC_HL_ZP18_S1_HL_ZP18_S2_Bayer2x2_ZONE2D", "FLG_DKH_FDL_Z2DorOF_FDL_Z2DorOF_ZL1not")
        Call ClearALLFlagBit("FLG_DKH_FDL_Z2DorOF_FDL_Z2DorOF_ZL1not")
        Call ClearALLFlagBit("FLG_DK_ZL1IC_HL_ZP18_S1_HL_ZP18_S2_Bayer2x2_ZONE2D")

    ' 5.FlagCopy.HLD_33FSC
    Dim sPlane28 As CImgPlane
    Call GetFreePlane(sPlane28, "Normal_Bayer2x2", idpDepthS16, True, "sPlane28")
    Call FlagCopy(sPlane28, "Bayer2x2_ZONE2D", "Flg_Temp1", 1)
        Call ClearALLFlagBit("Flg_Temp1")

    ' 6.スイッチ分岐.HLD_33FSC
    Dim Skip_HLD_33FSC_swid0 As Boolean
    Skip_HLD_33FSC_swid0 = False
    If tmp97 >= 2 Then
        Skip_HLD_33FSC_swid0 = True

        ' 8.Convolution.HLD_33FSC
        Dim sPlane29 As CImgPlane
        Call GetFreePlane(sPlane29, "Normal_Bayer2x2", idpDepthS16, , "sPlane29")
        Call Convolution(sPlane28, "Bayer2x2_ZONE2D", EEE_COLOR_ALL, sPlane29, "Bayer2x2_ZONE2D", EEE_COLOR_ALL, "kernel_Couplet")

        ' 9.CountBitMask.HLD_33FSC
        Dim tmp98_0(nSite) As Double
Call ClearALLFlagBit("Flg_Temp2")
        Call countBitMask(sPlane29, "Bayer2x2_ZONE2D", EEE_COLOR_ALL, idpCountAbove, 2, 2, idpLimitInclude, tmp98_0, "Flg_Temp2", , &HFFF8)

        ' 10.readPixelSite.HLD_33FSC
        Dim tmp_RPD9_0(nSite) As CPixInfo
        Call ReadPixelSite(sPlane29, "Bayer2x2_ZONE2D", tmp98_0, "Flg_Temp2", tmp_RPD9_0, idpAddrAbsolute)
        Dim tmp_RPD10(nSite) As CPixInfo
        Call RPDUnion(tmp_RPD10, tmp_RPD9_0)

        ' 12.MakeOtp.HLD_33FSC
        Dim tmp4_Info_Hadd_HLD_33FSC() As Double
        Dim tmp4_Info_Vadd_HLD_33FSC() As Double
        Dim tmp4_Info_Dire_HLD_33FSC() As Double
        Dim tmp4_Info_Sorc_HLD_33FSC() As Double
        Dim tmp4_Info_Count_HLD_33FSC(nSite) As Double
        Dim tmp99 As Double
'        Dim i As Long
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If tmp99 < tmp_RPD10(site).Count Then
                    tmp99 = tmp_RPD10(site).Count
                End If
            End If
        Next site
        ReDim tmp4_Info_Hadd_HLD_33FSC(nSite, tmp99) As Double
        ReDim tmp4_Info_Vadd_HLD_33FSC(nSite, tmp99) As Double
        ReDim tmp4_Info_Dire_HLD_33FSC(nSite, tmp99) As Double
        ReDim tmp4_Info_Sorc_HLD_33FSC(nSite, tmp99) As Double
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                For i = 0 To (tmp_RPD10(site).Count - 1)
                    tmp4_Info_Hadd_HLD_33FSC(site, i) = tmp_RPD10(site).PixInfo(i).x * 1 + (0)
                    tmp4_Info_Vadd_HLD_33FSC(site, i) = tmp_RPD10(site).PixInfo(i).y * 1 + (0)
                    tmp4_Info_Dire_HLD_33FSC(site, i) = 0
                    tmp4_Info_Sorc_HLD_33FSC(site, i) = 2
                Next i
                tmp4_Info_Count_HLD_33FSC(site) = tmp_RPD10(site).Count
            End If
        Next site

        ' 13.CountBitMask.HLD_33FSC
        Dim tmp100_0(nSite) As Double
        Call countBitMask(sPlane29, "Bayer2x2_ZONE2D", EEE_COLOR_ALL, idpCountAbove, 16, 16, idpLimitInclude, tmp100_0, "Flg_Temp3", , &HFFC7)

        ' 14.readPixelSite.HLD_33FSC
        Dim tmp_RPD11_0(nSite) As CPixInfo
        Call ReadPixelSite(sPlane29, "Bayer2x2_ZONE2D", tmp100_0, "Flg_Temp3", tmp_RPD11_0, idpAddrAbsolute)
        Dim tmp_RPD12(nSite) As CPixInfo
        Call RPDUnion(tmp_RPD12, tmp_RPD11_0)

        ' 16.MakeOtp.HLD_33FSC
        Dim tmp7_Info_Hadd_HLD_33FSC() As Double
        Dim tmp7_Info_Vadd_HLD_33FSC() As Double
        Dim tmp7_Info_Dire_HLD_33FSC() As Double
        Dim tmp7_Info_Sorc_HLD_33FSC() As Double
        Dim tmp7_Info_Count_HLD_33FSC(nSite) As Double
        Dim tmp101 As Double
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If tmp101 < tmp_RPD12(site).Count Then
                    tmp101 = tmp_RPD12(site).Count
                End If
            End If
        Next site
        ReDim tmp7_Info_Hadd_HLD_33FSC(nSite, tmp101) As Double
        ReDim tmp7_Info_Vadd_HLD_33FSC(nSite, tmp101) As Double
        ReDim tmp7_Info_Dire_HLD_33FSC(nSite, tmp101) As Double
        ReDim tmp7_Info_Sorc_HLD_33FSC(nSite, tmp101) As Double
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                For i = 0 To (tmp_RPD12(site).Count - 1)
                    tmp7_Info_Hadd_HLD_33FSC(site, i) = tmp_RPD12(site).PixInfo(i).x * 1 + (0)
                    tmp7_Info_Vadd_HLD_33FSC(site, i) = tmp_RPD12(site).PixInfo(i).y * 1 + (0)
                    tmp7_Info_Dire_HLD_33FSC(site, i) = 2
                    tmp7_Info_Sorc_HLD_33FSC(site, i) = 2
                Next i
                tmp7_Info_Count_HLD_33FSC(site) = tmp_RPD12(site).Count
            End If
        Next site

        ' 17.AddDefect.HLD_33FSC
        Dim tmp102(nSite) As Double
        Call GetSum(tmp102, tmp4_Info_Count_HLD_33FSC, tmp7_Info_Count_HLD_33FSC)
        Dim tmp103 As Double
        Dim tmp104 As Double
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                tmp103 = tmp102(site)
                If tmp103 > tmp104 Then
                    tmp104 = tmp102(site)
                End If
            End If
        Next site
        Dim tmp13_Info_Hadd_HLD_33FSC() As Double
        Dim tmp13_Info_Vadd_HLD_33FSC() As Double
        Dim tmp13_Info_Dire_HLD_33FSC() As Double
        Dim tmp13_Info_Sorc_HLD_33FSC() As Double
        Dim tmp13_Info_Count_HLD_33FSC(nSite) As Double
        ReDim tmp13_Info_Hadd_HLD_33FSC(nSite, tmp104) As Double
        ReDim tmp13_Info_Vadd_HLD_33FSC(nSite, tmp104) As Double
        ReDim tmp13_Info_Dire_HLD_33FSC(nSite, tmp104) As Double
        ReDim tmp13_Info_Sorc_HLD_33FSC(nSite, tmp104) As Double
        Dim tmp105(nSite) As Double
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                For i = 0 To tmp4_Info_Count_HLD_33FSC(site) - 1
                    tmp105(site) = tmp105(site) + 1
                    tmp13_Info_Hadd_HLD_33FSC(site, tmp105(site) - 1) = tmp4_Info_Hadd_HLD_33FSC(site, i)
                    tmp13_Info_Vadd_HLD_33FSC(site, tmp105(site) - 1) = tmp4_Info_Vadd_HLD_33FSC(site, i)
                    tmp13_Info_Dire_HLD_33FSC(site, tmp105(site) - 1) = tmp4_Info_Dire_HLD_33FSC(site, i)
                    tmp13_Info_Sorc_HLD_33FSC(site, tmp105(site) - 1) = tmp4_Info_Sorc_HLD_33FSC(site, i)
                Next i
            End If
        Next site
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                For i = 0 To tmp7_Info_Count_HLD_33FSC(site) - 1
                    tmp13_Info_Hadd_HLD_33FSC(site, tmp105(site) + i) = tmp7_Info_Hadd_HLD_33FSC(site, i)
                    tmp13_Info_Vadd_HLD_33FSC(site, tmp105(site) + i) = tmp7_Info_Vadd_HLD_33FSC(site, i)
                    tmp13_Info_Dire_HLD_33FSC(site, tmp105(site) + i) = tmp7_Info_Dire_HLD_33FSC(site, i)
                    tmp13_Info_Sorc_HLD_33FSC(site, tmp105(site) + i) = tmp7_Info_Sorc_HLD_33FSC(site, i)
                Next i
            End If
            tmp13_Info_Count_HLD_33FSC(site) = tmp102(site)
        Next site

        ' 18.CountBitMask.HLD_33FSC
        Dim tmp106_0(nSite) As Double
        Call countBitMask(sPlane29, "Bayer2x2_ZONE2D", EEE_COLOR_ALL, idpCountAbove, 128, 128, idpLimitInclude, tmp106_0, "Flg_Temp4", , &HFE3F)

        ' 19.readPixelSite.HLD_33FSC
        Dim tmp_RPD13_0(nSite) As CPixInfo
        Call ReadPixelSite(sPlane29, "Bayer2x2_ZONE2D", tmp106_0, "Flg_Temp4", tmp_RPD13_0, idpAddrAbsolute)
        Dim tmp_RPD14(nSite) As CPixInfo
        Call RPDUnion(tmp_RPD14, tmp_RPD13_0)

        ' 20.MakeOtp.HLD_33FSC
        Dim tmp15_Info_Hadd_HLD_33FSC() As Double
        Dim tmp15_Info_Vadd_HLD_33FSC() As Double
        Dim tmp15_Info_Dire_HLD_33FSC() As Double
        Dim tmp15_Info_Sorc_HLD_33FSC() As Double
        Dim tmp15_Info_Count_HLD_33FSC(nSite) As Double
        Dim tmp107 As Double
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If tmp107 < tmp_RPD14(site).Count Then
                    tmp107 = tmp_RPD14(site).Count
                End If
            End If
        Next site
        ReDim tmp15_Info_Hadd_HLD_33FSC(nSite, tmp107) As Double
        ReDim tmp15_Info_Vadd_HLD_33FSC(nSite, tmp107) As Double
        ReDim tmp15_Info_Dire_HLD_33FSC(nSite, tmp107) As Double
        ReDim tmp15_Info_Sorc_HLD_33FSC(nSite, tmp107) As Double
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                For i = 0 To (tmp_RPD14(site).Count - 1)
                    tmp15_Info_Hadd_HLD_33FSC(site, i) = tmp_RPD14(site).PixInfo(i).x * 1 + (0)
                    tmp15_Info_Vadd_HLD_33FSC(site, i) = tmp_RPD14(site).PixInfo(i).y * 1 + (0)
                    tmp15_Info_Dire_HLD_33FSC(site, i) = 1
                    tmp15_Info_Sorc_HLD_33FSC(site, i) = 2
                Next i
                tmp15_Info_Count_HLD_33FSC(site) = tmp_RPD14(site).Count
            End If
        Next site

        ' 21.AddDefect.HLD_33FSC
        Dim tmp108(nSite) As Double
        Call GetSum(tmp108, tmp13_Info_Count_HLD_33FSC, tmp15_Info_Count_HLD_33FSC)
        Dim tmp109 As Double
        Dim tmp110 As Double
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                tmp109 = tmp108(site)
                If tmp109 > tmp110 Then
                    tmp110 = tmp108(site)
                End If
            End If
        Next site
        Dim tmp21_Info_Hadd_HLD_33FSC() As Double
        Dim tmp21_Info_Vadd_HLD_33FSC() As Double
        Dim tmp21_Info_Dire_HLD_33FSC() As Double
        Dim tmp21_Info_Sorc_HLD_33FSC() As Double
        Dim tmp21_Info_Count_HLD_33FSC(nSite) As Double
        ReDim tmp21_Info_Hadd_HLD_33FSC(nSite, tmp110) As Double
        ReDim tmp21_Info_Vadd_HLD_33FSC(nSite, tmp110) As Double
        ReDim tmp21_Info_Dire_HLD_33FSC(nSite, tmp110) As Double
        ReDim tmp21_Info_Sorc_HLD_33FSC(nSite, tmp110) As Double
        Dim tmp111(nSite) As Double
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                For i = 0 To tmp13_Info_Count_HLD_33FSC(site) - 1
                    tmp111(site) = tmp111(site) + 1
                    tmp21_Info_Hadd_HLD_33FSC(site, tmp111(site) - 1) = tmp13_Info_Hadd_HLD_33FSC(site, i)
                    tmp21_Info_Vadd_HLD_33FSC(site, tmp111(site) - 1) = tmp13_Info_Vadd_HLD_33FSC(site, i)
                    tmp21_Info_Dire_HLD_33FSC(site, tmp111(site) - 1) = tmp13_Info_Dire_HLD_33FSC(site, i)
                    tmp21_Info_Sorc_HLD_33FSC(site, tmp111(site) - 1) = tmp13_Info_Sorc_HLD_33FSC(site, i)
                Next i
            End If
        Next site
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                For i = 0 To tmp15_Info_Count_HLD_33FSC(site) - 1
                    tmp21_Info_Hadd_HLD_33FSC(site, tmp111(site) + i) = tmp15_Info_Hadd_HLD_33FSC(site, i)
                    tmp21_Info_Vadd_HLD_33FSC(site, tmp111(site) + i) = tmp15_Info_Vadd_HLD_33FSC(site, i)
                    tmp21_Info_Dire_HLD_33FSC(site, tmp111(site) + i) = tmp15_Info_Dire_HLD_33FSC(site, i)
                    tmp21_Info_Sorc_HLD_33FSC(site, tmp111(site) + i) = tmp15_Info_Sorc_HLD_33FSC(site, i)
                Next i
            End If
            tmp21_Info_Count_HLD_33FSC(site) = tmp108(site)
        Next site

        ' 23.CountBitMask.HLD_33FSC
        Dim tmp112_0(nSite) As Double
        Call countBitMask(sPlane29, "Bayer2x2_ZONE2D", EEE_COLOR_ALL, idpCountAbove, 1024, 1024, idpLimitInclude, tmp112_0, "Flg_Temp5", , &HF1FF)

        ' 24.計算式評価.HLD_33FSC
        Dim tmp113(nSite) As Double
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                tmp113(site) = tmp98_0(site) + tmp100_0(site) + tmp106_0(site) + tmp112_0(site)
            End If
        Next site

        ' 25.readPixelSite.HLD_33FSC
        Dim tmp_RPD15_0(nSite) As CPixInfo
        Call ReadPixelSite(sPlane29, "Bayer2x2_ZONE2D", tmp112_0, "Flg_Temp5", tmp_RPD15_0, idpAddrAbsolute)
        Dim tmp_RPD16(nSite) As CPixInfo
        Call RPDUnion(tmp_RPD16, tmp_RPD15_0)

        ' 26.MakeOtp.HLD_33FSC
        Dim tmp26_Info_Hadd_HLD_33FSC() As Double
        Dim tmp26_Info_Vadd_HLD_33FSC() As Double
        Dim tmp26_Info_Dire_HLD_33FSC() As Double
        Dim tmp26_Info_Sorc_HLD_33FSC() As Double
        Dim tmp26_Info_Count_HLD_33FSC(nSite) As Double
        Dim tmp114 As Double
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If tmp114 < tmp_RPD16(site).Count Then
                    tmp114 = tmp_RPD16(site).Count
                End If
            End If
        Next site
        ReDim tmp26_Info_Hadd_HLD_33FSC(nSite, tmp114) As Double
        ReDim tmp26_Info_Vadd_HLD_33FSC(nSite, tmp114) As Double
        ReDim tmp26_Info_Dire_HLD_33FSC(nSite, tmp114) As Double
        ReDim tmp26_Info_Sorc_HLD_33FSC(nSite, tmp114) As Double
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                For i = 0 To (tmp_RPD16(site).Count - 1)
                    tmp26_Info_Hadd_HLD_33FSC(site, i) = tmp_RPD16(site).PixInfo(i).x * 1 + (0)
                    tmp26_Info_Vadd_HLD_33FSC(site, i) = tmp_RPD16(site).PixInfo(i).y * 1 + (0)
                    tmp26_Info_Dire_HLD_33FSC(site, i) = 3
                    tmp26_Info_Sorc_HLD_33FSC(site, i) = 2
                Next i
                tmp26_Info_Count_HLD_33FSC(site) = tmp_RPD16(site).Count
            End If
        Next site

        ' 27.AddDefect.HLD_33FSC
        Dim tmp115(nSite) As Double
        Call GetSum(tmp115, tmp21_Info_Count_HLD_33FSC, tmp26_Info_Count_HLD_33FSC)
        Dim tmp116 As Double
        Dim tmp117 As Double
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                tmp116 = tmp115(site)
                If tmp116 > tmp117 Then
                    tmp117 = tmp115(site)
                End If
            End If
        Next site
        Dim tmp32_Info_Hadd_HLD_33FSC() As Double
        Dim tmp32_Info_Vadd_HLD_33FSC() As Double
        Dim tmp32_Info_Dire_HLD_33FSC() As Double
        Dim tmp32_Info_Sorc_HLD_33FSC() As Double
        Dim tmp32_Info_Count_HLD_33FSC(nSite) As Double
        ReDim tmp32_Info_Hadd_HLD_33FSC(nSite, tmp117) As Double
        ReDim tmp32_Info_Vadd_HLD_33FSC(nSite, tmp117) As Double
        ReDim tmp32_Info_Dire_HLD_33FSC(nSite, tmp117) As Double
        ReDim tmp32_Info_Sorc_HLD_33FSC(nSite, tmp117) As Double
        Dim tmp118(nSite) As Double
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                For i = 0 To tmp21_Info_Count_HLD_33FSC(site) - 1
                    tmp118(site) = tmp118(site) + 1
                    tmp32_Info_Hadd_HLD_33FSC(site, tmp118(site) - 1) = tmp21_Info_Hadd_HLD_33FSC(site, i)
                    tmp32_Info_Vadd_HLD_33FSC(site, tmp118(site) - 1) = tmp21_Info_Vadd_HLD_33FSC(site, i)
                    tmp32_Info_Dire_HLD_33FSC(site, tmp118(site) - 1) = tmp21_Info_Dire_HLD_33FSC(site, i)
                    tmp32_Info_Sorc_HLD_33FSC(site, tmp118(site) - 1) = tmp21_Info_Sorc_HLD_33FSC(site, i)
                Next i
            End If
        Next site
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                For i = 0 To tmp26_Info_Count_HLD_33FSC(site) - 1
                    tmp32_Info_Hadd_HLD_33FSC(site, tmp118(site) + i) = tmp26_Info_Hadd_HLD_33FSC(site, i)
                    tmp32_Info_Vadd_HLD_33FSC(site, tmp118(site) + i) = tmp26_Info_Vadd_HLD_33FSC(site, i)
                    tmp32_Info_Dire_HLD_33FSC(site, tmp118(site) + i) = tmp26_Info_Dire_HLD_33FSC(site, i)
                    tmp32_Info_Sorc_HLD_33FSC(site, tmp118(site) + i) = tmp26_Info_Sorc_HLD_33FSC(site, i)
                Next i
            End If
            tmp32_Info_Count_HLD_33FSC(site) = tmp115(site)
        Next site

        ' 28.PutDefectResult.HLD_33FSC
        Call ResultAdd("HLD_33FSC_Info_Num", tmp32_Info_Count_HLD_33FSC)
        Call ResultAdd("HLD_33FSC_Info_Hadd", tmp32_Info_Hadd_HLD_33FSC)
        Call ResultAdd("HLD_33FSC_Info_Vadd", tmp32_Info_Vadd_HLD_33FSC)
        Call ResultAdd("HLD_33FSC_Info_Dire", tmp32_Info_Dire_HLD_33FSC)
        Call ResultAdd("HLD_33FSC_Info_Sorc", tmp32_Info_Sorc_HLD_33FSC)

        ' 30.RPDUnion.HLD_33FSC
        Dim tmp_RPD17(nSite) As CPixInfo
        Call RPDUnion(tmp_RPD17, tmp_RPD10, tmp_RPD12, tmp_RPD14, tmp_RPD16)

        ' 31.WritePixelAddrSite.HLD_33FSC
        Dim sPlane30 As CImgPlane
        Call GetFreePlane(sPlane30, "Normal_Bayer2x2", idpDepthS16, True, "sPlane30")
        Call WritePixelAddrSite(sPlane30, "Bayer2x2_FULL", tmp_RPD17)

        ' 32.PutFlag_FA.HLD_33FSC
        Call PutFlag_FA(sPlane30, "Bayer2x2_FULL", EEE_COLOR_ALL, idpCountAbove, 1, 1, idpLimitEachSite, idpLimitInclude, "FLG_HLD_33FSC")

        ' 0-Else.処理分岐(else).HLD_33FSC
    Else

        ' 33.連続点Skip処理.HLD_33FSC
        Dim tmp119(nSite) As Double
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                tmp119(site) = 0
            End If
        Next site
        Call TheIDP.PlaneManager("Normal_Bayer2x2").GetSharedFlagPlane("FLG_HLD_33FSC").SetFlagBit("FLG_HLD_33FSC")
        Dim tmp120() As Double
        ReDim tmp120(nSite, 0) As Double
        Call ResultAdd("HLD_33FSC_Info_Num", tmp119)
        Call ResultAdd("HLD_33FSC_Info_Hadd", tmp120)
        Call ResultAdd("HLD_33FSC_Info_Vadd", tmp120)
        Call ResultAdd("HLD_33FSC_Info_Dire", tmp120)
        Call ResultAdd("HLD_33FSC_Info_Sorc", tmp120)

        ' 34.合流N.HLD_33FSC
    End If
        Call ClearALLFlagBit("FLG_HLD_33FSC")
        Call ClearALLFlagBit("Flg_Temp5")
        Call ClearALLFlagBit("Flg_Temp4")
        Call ClearALLFlagBit("Flg_Temp3")
        Call ClearALLFlagBit("Flg_Temp2")
        Call ReleasePlane(sPlane28)
        Call ReleasePlane(sPlane29)
        Call ReleasePlane(sPlane30)
    Dim tmp121() As Double
    tmp121 = IIf(Skip_HLD_33FSC_swid0, tmp113, tmp119)

    ' 35.PutTestResult.HLD_33FSC
    Call ResultAdd("HLD_33FSC", tmp121)

' #### RD_HLDFD ####

    ' 0.項目和.RD_HLDFD
    Dim tmp_HLD_33FSC() As Double
    TheResult.GetResult "HLD_33FSC", tmp_HLD_33FSC
    Dim tmp122(nSite) As Double
    Call GetSum(tmp122, tmp_DKH_FDL_Z2D, tmp_HLD_33FSC, tmp_OF_FDL_Z2D, tmp_OF_ZL1)
    Dim tmp123(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            tmp123(site) = 3 + 1
        End If
    Next site

    ' 1.PutTestResult.RD_HLDFD
    Call ResultAdd("RD_HLDFD", tmp122)

End Function



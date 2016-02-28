Attribute VB_Name = "Image_005_LL_ERR_Mod"

Option Explicit

Public Function LL_ERR_Process()

        Call PutImageInto_Common

' #### LL_SEN ####

    Dim site As Long

    ' 0.画像情報インポート.LL_SEN
    Dim LL_ERR_Param As CParamPlane
    Dim LL_ERR_DevInfo As CDeviceConfigInfo
    Dim LL_ERR_Plane As CImgPlane
    Set LL_ERR_Param = TheParameterBank.Item("LLImageTest_Acq1")
    Set LL_ERR_DevInfo = TheDeviceProfiler.ConfigInfo("LLImageTest_Acq1")
    Set LL_ERR_Plane = LL_ERR_Param.plane

    ' 1.Clamp.LL_SEN
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call Clamp(LL_ERR_Plane, sPlane1, "Bayer2x4_VOPB")

    ' 2.Median.LL_SEN
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 1, 5)

    ' 3.Median.LL_SEN
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call MedianEx(sPlane2, sPlane3, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane2)

    ' 82.Average_FA.LL_SEN
    Dim tmp1_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, tmp1_0)
    Dim tmp2 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp2, tmp1_0)

    ' 83.GetAverage_Color.LL_SEN
    Dim tmp3(nSite) As Double
    Call GetAverage_Color(tmp3, tmp2, "Gr1", "Gb1", "Gr2", "Gb2")

    ' 233.LSB定義.LL_SEN
    Dim LL_ERR_LSB() As Double
     LL_ERR_LSB = LL_ERR_DevInfo.Lsb.AsDouble

    ' 238.LSB換算.LL_SEN
    Dim tmp4(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp4(site) = tmp3(site) * LL_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.LL_SEN
    Call ResultAdd("LL_SEN", tmp4)

' #### LL_BZL0 ####

    ' 0.画像情報インポート.LL_BZL0

    ' 1.Clamp.LL_BZL0

    ' 2.Median.LL_BZL0

    ' 3.Median.LL_BZL0

    ' 4.Subtract(通常).LL_BZL0
    Dim sPlane4 As CImgPlane
    Call GetFreePlane(sPlane4, "Normal_Bayer2x4", idpDepthS16, , "sPlane4")
    Call Subtract(sPlane1, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane3, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane4, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane3)

    ' 5.LSB定義.LL_BZL0

    ' 6.SliceLevel生成.LL_BZL0
    Dim tmp_Slice1(nSite) As Double
    Call MakeSliceLevel(tmp_Slice1, -0.0008, LL_ERR_LSB, , , , idpCountBelow)

    ' 7.マスク取得.LL_BZL0

    Call SharedFlagNot("Normal_Bayer2x4", "Bayer2x4_ZONE2D", "FLG_DK_ZL1ICnot", "FLG_DK_ZL1IC")
        Call ClearALLFlagBit("FLG_DK_ZL1IC")

    ' 8.Count_FA.LL_BZL0
    Dim tmp5_0 As CImgColorAllResult
    Call count_FA(sPlane4, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountBelow, tmp_Slice1, tmp_Slice1, idpLimitEachSite, idpLimitExclude, tmp5_0, "FLG_LL_BZL0", "FLG_DK_ZL1ICnot")
        Call ClearALLFlagBit("FLG_LL_BZL0")
    Dim tmp6 As CImgColorAllResult
    Call GetSum_CImgColor(tmp6, tmp5_0)

    ' 9.GetSum_Color.LL_BZL0
    Dim tmp7(nSite) As Double
    Call GetSum_Color(tmp7, tmp6, "-")

    ' 10.PutTestResult.LL_BZL0
    Call ResultAdd("LL_BZL0", tmp7)

' #### LL_BZL1 ####

    ' 0.画像情報インポート.LL_BZL1

    ' 1.Clamp.LL_BZL1

    ' 2.Median.LL_BZL1

    ' 3.Median.LL_BZL1

    ' 4.Subtract(通常).LL_BZL1

    ' 5.LSB定義.LL_BZL1

    ' 6.SliceLevel生成.LL_BZL1
    Dim tmp_Slice2(nSite) As Double
    Call MakeSliceLevel(tmp_Slice2, -0.0013, LL_ERR_LSB, , , , idpCountBelow)

    ' 7.マスク取得.LL_BZL1

    ' 8.Count_FA.LL_BZL1
    Dim tmp8_0 As CImgColorAllResult
    Call count_FA(sPlane4, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountBelow, tmp_Slice2, tmp_Slice2, idpLimitEachSite, idpLimitExclude, tmp8_0, , "FLG_DK_ZL1ICnot")
    Dim tmp9 As CImgColorAllResult
    Call GetSum_CImgColor(tmp9, tmp8_0)

    ' 9.GetSum_Color.LL_BZL1
    Dim tmp10(nSite) As Double
    Call GetSum_Color(tmp10, tmp9, "-")

    ' 10.PutTestResult.LL_BZL1
    Call ResultAdd("LL_BZL1", tmp10)

    ' 14.d_read_vmcu_point.LL_BZL1
    If (Flg_Debug = 1 Or Sw_Ana = 1) Then
        Call d_read_vmcu_Point(sPlane4, "Bayer2x4_ZONE2D", 500, LL_ERR_LSB, "NoKCO", "LLB", "mV", idpCountBelow, tmp_Slice2, idpLimitExclude, "FLG_DK_ZL1ICnot", "-")
    End If

' #### LL_WZL0 ####

    ' 0.画像情報インポート.LL_WZL0

    ' 1.Clamp.LL_WZL0

    ' 2.Median.LL_WZL0

    ' 3.Median.LL_WZL0

    ' 4.Subtract(通常).LL_WZL0

    ' 5.LSB定義.LL_WZL0

    ' 6.SliceLevel生成.LL_WZL0
    Dim tmp_Slice3(nSite) As Double
    Call MakeSliceLevel(tmp_Slice3, 0.0008, LL_ERR_LSB, , , , idpCountAbove)

    ' 7.マスク取得.LL_WZL0

    ' 8.Count_FA.LL_WZL0
    Dim tmp11_0 As CImgColorAllResult
    Call count_FA(sPlane4, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice3, tmp_Slice3, idpLimitEachSite, idpLimitExclude, tmp11_0, "FLG_LL_WZL0", "FLG_DK_ZL1ICnot")
        Call ClearALLFlagBit("FLG_LL_WZL0")
    Dim tmp12 As CImgColorAllResult
    Call GetSum_CImgColor(tmp12, tmp11_0)

    ' 9.GetSum_Color.LL_WZL0
    Dim tmp13(nSite) As Double
    Call GetSum_Color(tmp13, tmp12, "-")

    ' 10.PutTestResult.LL_WZL0
    Call ResultAdd("LL_WZL0", tmp13)

' #### LL_WZL1 ####

    ' 0.画像情報インポート.LL_WZL1

    ' 1.Clamp.LL_WZL1

    ' 2.Median.LL_WZL1

    ' 3.Median.LL_WZL1

    ' 4.Subtract(通常).LL_WZL1

    ' 5.LSB定義.LL_WZL1

    ' 6.SliceLevel生成.LL_WZL1
    Dim tmp_Slice4(nSite) As Double
    Call MakeSliceLevel(tmp_Slice4, 0.0013, LL_ERR_LSB, , , , idpCountAbove)

    ' 7.マスク取得.LL_WZL1

    ' 8.Count_FA.LL_WZL1
    Dim tmp14_0 As CImgColorAllResult
    Call count_FA(sPlane4, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice4, tmp_Slice4, idpLimitEachSite, idpLimitExclude, tmp14_0, , "FLG_DK_ZL1ICnot")
    Dim tmp15 As CImgColorAllResult
    Call GetSum_CImgColor(tmp15, tmp14_0)

    ' 9.GetSum_Color.LL_WZL1
    Dim tmp16(nSite) As Double
    Call GetSum_Color(tmp16, tmp15, "-")

    ' 10.PutTestResult.LL_WZL1
    Call ResultAdd("LL_WZL1", tmp16)

    ' 14.d_read_vmcu_point.LL_WZL1
    If (Flg_Debug = 1 Or Sw_Ana = 1) Then
        Call d_read_vmcu_Point(sPlane4, "Bayer2x4_ZONE2D", 500, LL_ERR_LSB, "NoKCO", "LLW", "mV", idpCountAbove, tmp_Slice4, idpLimitExclude, "FLG_DK_ZL1ICnot", "-")
    End If
        Call ClearALLFlagBit("FLG_DK_ZL1ICnot")
        Call ReleasePlane(sPlane4)

' #### LL_HLN ####

    ' 0.画像情報インポート.LL_HLN

    ' 1.Clamp.LL_HLN

    ' 2.Median.LL_HLN
    Dim sPlane5 As CImgPlane
    Call GetFreePlane(sPlane5, "Normal_Bayer2x4", idpDepthS16, , "sPlane5")
    Call MedianEx(sPlane1, sPlane5, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane1)

    ' 3.ZONE取得.LL_HLN

    ' 5.AccumulateRow.LL_HLN
    Dim sPlane6 As CImgPlane
    Call GetFreePlane(sPlane6, "Normal_Bayer2x4_ACR", idpDepthF32, , "sPlane6")
    Call MakeAcrPMD(sPlane6, "Bayer2x4_ZONE2D", "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT)
    Call AccumulateRow(sPlane5, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane6, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, idpAccumMean)
        Call ReleasePlane(sPlane5)

    ' 6.SubRows.LL_HLN
    Dim sPlane7 As CImgPlane
    Call GetFreePlane(sPlane7, "Normal_Bayer2x4_ACR", idpDepthF32, , "sPlane7")
    Call SubRows(sPlane6, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, sPlane7, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, 2)
        Call ReleasePlane(sPlane6)
    Call MakeAcrJudgePMD(sPlane7, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", "Bayer2x4_ACR_2_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, 2)

    ' 9.AbsMax_FA.LL_HLN
    Dim tmp17_0 As CImgColorAllResult
    Call AbsMax_FA(sPlane7, "Bayer2x4_ACR_2_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_ALL, tmp17_0)
    Dim tmp18 As CImgColorAllResult
    Call GetMax_CImgColor(tmp18, tmp17_0)

    ' 10.GetAbsMax_Color.LL_HLN
    Dim tmp19(nSite) As Double
    Call GetAbsMax_Color(tmp19, tmp18, "-")

    ' 13.GetAbs.LL_HLN
    Dim tmp20(nSite) As Double
    Call GetAbs(tmp20, tmp19)

    ' 14.パラメータ取得.LL_HLN
    Dim tmp_LL_SEN() As Double
    TheResult.GetResult "LL_SEN", tmp_LL_SEN
    LL_ERR_LSB = TheDeviceProfiler.ConfigInfo("LLImageTest_Acq1").Lsb.AsDouble
        Call TheParameterBank.Delete("LLImageTest_Acq1")
    Dim tmp21(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp21(site) = Div(tmp_LL_SEN(site), LL_ERR_LSB(site), 0)
        End If
    Next site

    ' 15.計算式評価.LL_HLN
    Dim tmp22(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp22(site) = Div(tmp20(site), tmp21(site), 999)
        End If
    Next site

    ' 16.d_read_vmcu_Line.LL_HLN
    If (Flg_Debug = 1 Or Sw_Ana = 1) Then
        Call d_read_vmcu_Line(sPlane7, "Bayer2x4_ACR_2_ZONE2D_EEE_COLOR_FLAT", 1, LL_ERR_LSB, "NoKCO", "LL_HLINE", "%", tmp19, "HLINE", "ABSMAX", tmp21)
    End If
        Call ReleasePlane(sPlane7)

    ' 17.LSB定義.LL_HLN

    ' 19.PutTestResult.LL_HLN
    Call ResultAdd("LL_HLN", tmp22)

End Function



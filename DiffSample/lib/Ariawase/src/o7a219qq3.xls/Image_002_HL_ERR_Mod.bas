Attribute VB_Name = "Image_002_HL_ERR_Mod"

Option Explicit

Public Function HL_ERR_Process()

        Call PutImageInto_Common

' #### HL_SEN ####

    Dim site As Long

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_SEN
    Dim HL_ERR_Param As CParamPlane
    Dim HL_ERR_DevInfo As CDeviceConfigInfo
    Dim HL_ERR_Plane As CImgPlane
    Set HL_ERR_Param = TheParameterBank.Item("HLImageTest_Acq1")
    Set HL_ERR_DevInfo = TheDeviceProfiler.ConfigInfo("HLImageTest_Acq1")
    Set HL_ERR_Plane = HL_ERR_Param.plane

    ' 1.Clamp.HL_SEN
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "HL_TEMP_Bayer2x4", idpDepthS16, , "sPlane1")
    Call Clamp(HL_ERR_Plane, sPlane1, "HL_TEMP_Bayer2x4_VOPB")

    ' 2.Median.HL_SEN
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "HL_TEMP_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "HL_TEMP_Bayer2x4_ZONE3", 1, 5)

    ' 3.Median.HL_SEN
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "HL_TEMP_Bayer2x4", idpDepthS16, , "sPlane3")
    Call MedianEx(sPlane2, sPlane3, "HL_TEMP_Bayer2x4_ZONE3", 5, 1)

    ' 82.Average_FA.HL_SEN
    Dim tmp1_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "HL_TEMP_Bayer2x4_ZONE0", EEE_COLOR_ALL, tmp1_0)
    Dim tmp2 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp2, tmp1_0)

    ' 83.GetAverage_Color.HL_SEN
    Dim tmp3(nSite) As Double
    Call GetAverage_Color(tmp3, tmp2, "Gr1", "Gb1", "Gr2", "Gb2")

    ' 233.LSBíËã`.HL_SEN
    Dim HL_ERR_LSB() As Double
     HL_ERR_LSB = HL_ERR_DevInfo.Lsb.AsDouble

    ' 238.LSBä∑éZ.HL_SEN
    Dim tmp4(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp4(site) = tmp3(site) * HL_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.HL_SEN
    Call ResultAdd("HL_SEN", tmp4)

' #### HL_HLN ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_HLN

    ' 1.Clamp.HL_HLN

    ' 2.Median.HL_HLN
    Dim sPlane4 As CImgPlane
    Call GetFreePlane(sPlane4, "HL_TEMP_Bayer2x4", idpDepthS16, , "sPlane4")
    Call MedianEx(sPlane1, sPlane4, "HL_TEMP_Bayer2x4_ZONE3", 5, 1)

    ' 3.ZONEéÊìæ.HL_HLN

    ' 5.AccumulateRow.HL_HLN
    Dim sPlane5 As CImgPlane
    Call GetFreePlane(sPlane5, "HL_TEMP_Bayer2x4_ACR", idpDepthF32, , "sPlane5")
    Call MakeAcrPMD(sPlane5, "HL_TEMP_Bayer2x4_ZONE2D", "HL_TEMP_Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT)
    Call AccumulateRow(sPlane4, "HL_TEMP_Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane5, "HL_TEMP_Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, idpAccumMean)
        Call ReleasePlane(sPlane4)

    ' 6.SubRows.HL_HLN
    Dim sPlane6 As CImgPlane
    Call GetFreePlane(sPlane6, "HL_TEMP_Bayer2x4_ACR", idpDepthF32, , "sPlane6")
    Call SubRows(sPlane5, "HL_TEMP_Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, sPlane6, "HL_TEMP_Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, 2)
        Call ReleasePlane(sPlane5)
    Call MakeAcrJudgePMD(sPlane6, "HL_TEMP_Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", "HL_TEMP_Bayer2x4_ACR_2_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, 2)

    ' 9.AbsMax_FA.HL_HLN
    Dim tmp5_0 As CImgColorAllResult
    Call AbsMax_FA(sPlane6, "HL_TEMP_Bayer2x4_ACR_2_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_ALL, tmp5_0)
    Dim tmp6 As CImgColorAllResult
    Call GetMax_CImgColor(tmp6, tmp5_0)

    ' 10.GetAbsMax_Color.HL_HLN
    Dim tmp7(nSite) As Double
    Call GetAbsMax_Color(tmp7, tmp6, "-")

    ' 13.GetAbs.HL_HLN
    Dim tmp8(nSite) As Double
    Call GetAbs(tmp8, tmp7)

    ' 14.ÉpÉâÉÅÅ[É^éÊìæ.HL_HLN
    Dim tmp_HL_SEN() As Double
    TheResult.GetResult "HL_SEN", tmp_HL_SEN
    HL_ERR_LSB = TheDeviceProfiler.ConfigInfo("HLImageTest_Acq1").Lsb.AsDouble
    Dim tmp9(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp9(site) = Div(tmp_HL_SEN(site), HL_ERR_LSB(site), 0)
        End If
    Next site

    ' 15.åvéZéÆï]âø.HL_HLN
    Dim tmp10(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp10(site) = Div(tmp8(site), tmp9(site), 999)
        End If
    Next site

    ' 16.d_read_vmcu_Line.HL_HLN
    If (Flg_Debug = 1 Or Sw_Ana = 1) Then
        Call d_read_vmcu_Line(sPlane6, "HL_TEMP_Bayer2x4_ACR_2_ZONE2D_EEE_COLOR_FLAT", 1, HL_ERR_LSB, "NoKCO", "HL_HLINE", "%", tmp7, "HLINE", "ABSMAX", tmp9)
    End If
        Call ReleasePlane(sPlane6)

    ' 17.LSBíËã`.HL_HLN

    ' 19.PutTestResult.HL_HLN
    Call ResultAdd("HL_HLN", tmp10)

' #### HL_VLN ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_VLN

    ' 1.Clamp.HL_VLN

    ' 2.Median.HL_VLN

    ' 3.ZONEéÊìæ.HL_VLN

    ' 5.AccumulateColumn.HL_VLN
    Dim sPlane7 As CImgPlane
    Call GetFreePlane(sPlane7, "HL_TEMP_Bayer2x4_ACC", idpDepthF32, , "sPlane7")
    Call MakeAccPMD(sPlane7, "HL_TEMP_Bayer2x4_ZONE2D", "HL_TEMP_Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT)
    Call AccumulateColumn(sPlane2, "HL_TEMP_Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane7, "HL_TEMP_Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, idpAccumMean)
        Call ReleasePlane(sPlane2)

    ' 6.SubColumns.HL_VLN
    Dim sPlane8 As CImgPlane
    Call GetFreePlane(sPlane8, "HL_TEMP_Bayer2x4_ACC", idpDepthF32, , "sPlane8")
    Call SubColumns(sPlane7, "HL_TEMP_Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, sPlane8, "HL_TEMP_Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, 2)
        Call ReleasePlane(sPlane7)
    Call MakeAccJudgePMD(sPlane8, "HL_TEMP_Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", "HL_TEMP_Bayer2x4_ACC_2_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, 2)

    ' 7.AbsMax_FA.HL_VLN
    Dim tmp11_0 As CImgColorAllResult
    Call AbsMax_FA(sPlane8, "HL_TEMP_Bayer2x4_ACC_2_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_ALL, tmp11_0)
    Dim tmp12 As CImgColorAllResult
    Call GetMax_CImgColor(tmp12, tmp11_0)

    ' 8.GetAbsMax_Color.HL_VLN
    Dim tmp13(nSite) As Double
    Call GetAbsMax_Color(tmp13, tmp12, "-")

    ' 11.GetAbs.HL_VLN
    Dim tmp14(nSite) As Double
    Call GetAbs(tmp14, tmp13)

    ' 14.ÉpÉâÉÅÅ[É^éÊìæ.HL_VLN

    ' 15.åvéZéÆï]âø.HL_VLN
    Dim tmp15(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp15(site) = Div(tmp14(site), tmp9(site), 999)
        End If
    Next site

    ' 16.d_read_vmcu_Line.HL_VLN
    If (Flg_Debug = 1 Or Sw_Ana = 1) Then
        Call d_read_vmcu_Line(sPlane8, "HL_TEMP_Bayer2x4_ACC_2_ZONE2D_EEE_COLOR_FLAT", 1, HL_ERR_LSB, "NoKCO", "HL_VLINE", "%", tmp13, "VLINE", "ABSMAX", tmp9)
    End If
        Call ReleasePlane(sPlane8)

    ' 17.LSBíËã`.HL_VLN

    ' 19.PutTestResult.HL_VLN
    Call ResultAdd("HL_VLN", tmp15)

' #### HL_SRRG ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_SRRG

    ' 1.Clamp.HL_SRRG

    ' 2.Median.HL_SRRG

    ' 3.Median.HL_SRRG

    ' 76.Average_FA.HL_SRRG

    ' 77.GetAverage_Color.HL_SRRG
    Dim tmp16(nSite) As Double
    Call GetAverage_Color(tmp16, tmp2, "Gr1", "Gr2")

    ' 78.GetAverage_Color.HL_SRRG
    Dim tmp17(nSite) As Double
    Call GetAverage_Color(tmp17, tmp2, "R1", "R2")

    ' 79.åvéZéÆï]âø.HL_SRRG
    Dim tmp18(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp18(site) = Div(tmp17(site), tmp16(site), 999)
        End If
    Next site

    ' 80.PutTestResult.HL_SRRG
    Call ResultAdd("HL_SRRG", tmp18)

' #### HL_SRBG ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_SRBG

    ' 1.Clamp.HL_SRBG

    ' 2.Median.HL_SRBG

    ' 3.Median.HL_SRBG

    ' 76.Average_FA.HL_SRBG

    ' 77.GetAverage_Color.HL_SRBG
    Dim tmp19(nSite) As Double
    Call GetAverage_Color(tmp19, tmp2, "Gb1", "Gb2")

    ' 78.GetAverage_Color.HL_SRBG
    Dim tmp20(nSite) As Double
    Call GetAverage_Color(tmp20, tmp2, "B1", "B2")

    ' 79.åvéZéÆï]âø.HL_SRBG
    Dim tmp21(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp21(site) = Div(tmp20(site), tmp19(site), 999)
        End If
    Next site

    ' 80.PutTestResult.HL_SRBG
    Call ResultAdd("HL_SRBG", tmp21)

' #### HL_SRGG ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_SRGG

    ' 1.Clamp.HL_SRGG

    ' 2.Median.HL_SRGG

    ' 3.Median.HL_SRGG

    ' 76.Average_FA.HL_SRGG

    ' 77.GetAverage_Color.HL_SRGG

    ' 78.GetAverage_Color.HL_SRGG

    ' 79.åvéZéÆï]âø.HL_SRGG
    Dim tmp22(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp22(site) = Div(tmp16(site), tmp19(site), 999)
        End If
    Next site

    ' 80.PutTestResult.HL_SRGG
    Call ResultAdd("HL_SRGG", tmp22)

' #### HL_CS3RG ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3RG

    ' 1.Clamp.HL_CS3RG

    ' 2.Median.HL_CS3RG

    ' 3.Median.HL_CS3RG

    ' 76.Average_FA.HL_CS3RG
    Dim tmp23_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "HL_TEMP_Bayer2x4_ZONEC2", EEE_COLOR_ALL, tmp23_0)
    Dim tmp24 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp24, tmp23_0)

    ' 77.GetAverage_Color.HL_CS3RG
    Dim tmp25(nSite) As Double
    Call GetAverage_Color(tmp25, tmp24, "Gr1", "Gb1", "Gr2", "Gb2")

    ' 78.GetAverage_Color.HL_CS3RG
    Dim tmp26(nSite) As Double
    Call GetAverage_Color(tmp26, tmp24, "R1", "R2")

    ' 79.åvéZéÆï]âø.HL_CS3RG
    Dim tmp27(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp27(site) = Div(tmp26(site), tmp25(site), 999)
        End If
    Next site

    ' 80.PutTestResult.HL_CS3RG
    Call ResultAdd("HL_CS3RG", tmp27)

' #### HL_CS3BG ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3BG

    ' 1.Clamp.HL_CS3BG

    ' 2.Median.HL_CS3BG

    ' 3.Median.HL_CS3BG

    ' 76.Average_FA.HL_CS3BG

    ' 77.GetAverage_Color.HL_CS3BG

    ' 78.GetAverage_Color.HL_CS3BG
    Dim tmp28(nSite) As Double
    Call GetAverage_Color(tmp28, tmp24, "B1", "B2")

    ' 79.åvéZéÆï]âø.HL_CS3BG
    Dim tmp29(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp29(site) = Div(tmp28(site), tmp25(site), 999)
        End If
    Next site

    ' 80.PutTestResult.HL_CS3BG
    Call ResultAdd("HL_CS3BG", tmp29)

' #### HL_CS3RG_ZL1 ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3RG_ZL1

    ' 1.Clamp.HL_CS3RG_ZL1

    ' 2.Median.HL_CS3RG_ZL1

    ' 3.Median.HL_CS3RG_ZL1

    ' 175.Average_FA.HL_CS3RG_ZL1

    ' 176.GetAverage_Color.HL_CS3RG_ZL1

    ' 177.GetAverage_Color.HL_CS3RG_ZL1

    ' 178.åvéZéÆï]âø.HL_CS3RG_ZL1

    ' 179.Average_FA.HL_CS3RG_ZL1
    Dim tmp30_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "HL_TEMP_Bayer2x4_ZONEL1", EEE_COLOR_ALL, tmp30_0)
    Dim tmp31 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp31, tmp30_0)

    ' 180.GetAverage_Color.HL_CS3RG_ZL1
    Dim tmp32(nSite) As Double
    Call GetAverage_Color(tmp32, tmp31, "Gr1", "Gb1", "Gr2", "Gb2")

    ' 181.GetAverage_Color.HL_CS3RG_ZL1
    Dim tmp33(nSite) As Double
    Call GetAverage_Color(tmp33, tmp31, "R1", "R2")

    ' 182.åvéZéÆï]âø.HL_CS3RG_ZL1
    Dim tmp34(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp34(site) = Div(tmp33(site), tmp32(site), 999)
        End If
    Next site

    ' 183.åvéZéÆï]âø.HL_CS3RG_ZL1
    Dim tmp35(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp35(site) = Div(tmp34(site), tmp27(site), 999) - 1
        End If
    Next site

    ' 184.PutTestResult.HL_CS3RG_ZL1
    Call ResultAdd("HL_CS3RG_ZL1", tmp35)

' #### HL_CS3RG_ZL2 ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3RG_ZL2

    ' 1.Clamp.HL_CS3RG_ZL2

    ' 2.Median.HL_CS3RG_ZL2

    ' 3.Median.HL_CS3RG_ZL2

    ' 175.Average_FA.HL_CS3RG_ZL2

    ' 176.GetAverage_Color.HL_CS3RG_ZL2

    ' 177.GetAverage_Color.HL_CS3RG_ZL2

    ' 178.åvéZéÆï]âø.HL_CS3RG_ZL2

    ' 179.Average_FA.HL_CS3RG_ZL2
    Dim tmp36_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "HL_TEMP_Bayer2x4_ZONEL2", EEE_COLOR_ALL, tmp36_0)
    Dim tmp37 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp37, tmp36_0)

    ' 180.GetAverage_Color.HL_CS3RG_ZL2
    Dim tmp38(nSite) As Double
    Call GetAverage_Color(tmp38, tmp37, "Gr1", "Gb1", "Gr2", "Gb2")

    ' 181.GetAverage_Color.HL_CS3RG_ZL2
    Dim tmp39(nSite) As Double
    Call GetAverage_Color(tmp39, tmp37, "R1", "R2")

    ' 182.åvéZéÆï]âø.HL_CS3RG_ZL2
    Dim tmp40(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp40(site) = Div(tmp39(site), tmp38(site), 999)
        End If
    Next site

    ' 183.åvéZéÆï]âø.HL_CS3RG_ZL2
    Dim tmp41(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp41(site) = Div(tmp40(site), tmp27(site), 999) - 1
        End If
    Next site

    ' 184.PutTestResult.HL_CS3RG_ZL2
    Call ResultAdd("HL_CS3RG_ZL2", tmp41)

' #### HL_CS3RG_ZL3 ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3RG_ZL3

    ' 1.Clamp.HL_CS3RG_ZL3

    ' 2.Median.HL_CS3RG_ZL3

    ' 3.Median.HL_CS3RG_ZL3

    ' 175.Average_FA.HL_CS3RG_ZL3

    ' 176.GetAverage_Color.HL_CS3RG_ZL3

    ' 177.GetAverage_Color.HL_CS3RG_ZL3

    ' 178.åvéZéÆï]âø.HL_CS3RG_ZL3

    ' 179.Average_FA.HL_CS3RG_ZL3
    Dim tmp42_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "HL_TEMP_Bayer2x4_ZONEL3", EEE_COLOR_ALL, tmp42_0)
    Dim tmp43 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp43, tmp42_0)

    ' 180.GetAverage_Color.HL_CS3RG_ZL3
    Dim tmp44(nSite) As Double
    Call GetAverage_Color(tmp44, tmp43, "Gr1", "Gb1", "Gr2", "Gb2")

    ' 181.GetAverage_Color.HL_CS3RG_ZL3
    Dim tmp45(nSite) As Double
    Call GetAverage_Color(tmp45, tmp43, "R1", "R2")

    ' 182.åvéZéÆï]âø.HL_CS3RG_ZL3
    Dim tmp46(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp46(site) = Div(tmp45(site), tmp44(site), 999)
        End If
    Next site

    ' 183.åvéZéÆï]âø.HL_CS3RG_ZL3
    Dim tmp47(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp47(site) = Div(tmp46(site), tmp27(site), 999) - 1
        End If
    Next site

    ' 184.PutTestResult.HL_CS3RG_ZL3
    Call ResultAdd("HL_CS3RG_ZL3", tmp47)

' #### HL_CS3RG_ZC1 ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3RG_ZC1

    ' 1.Clamp.HL_CS3RG_ZC1

    ' 2.Median.HL_CS3RG_ZC1

    ' 3.Median.HL_CS3RG_ZC1

    ' 175.Average_FA.HL_CS3RG_ZC1

    ' 176.GetAverage_Color.HL_CS3RG_ZC1

    ' 177.GetAverage_Color.HL_CS3RG_ZC1

    ' 178.åvéZéÆï]âø.HL_CS3RG_ZC1

    ' 179.Average_FA.HL_CS3RG_ZC1
    Dim tmp48_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "HL_TEMP_Bayer2x4_ZONEC1", EEE_COLOR_ALL, tmp48_0)
    Dim tmp49 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp49, tmp48_0)

    ' 180.GetAverage_Color.HL_CS3RG_ZC1
    Dim tmp50(nSite) As Double
    Call GetAverage_Color(tmp50, tmp49, "Gr1", "Gb1", "Gr2", "Gb2")

    ' 181.GetAverage_Color.HL_CS3RG_ZC1
    Dim tmp51(nSite) As Double
    Call GetAverage_Color(tmp51, tmp49, "R1", "R2")

    ' 182.åvéZéÆï]âø.HL_CS3RG_ZC1
    Dim tmp52(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp52(site) = Div(tmp51(site), tmp50(site), 999)
        End If
    Next site

    ' 183.åvéZéÆï]âø.HL_CS3RG_ZC1
    Dim tmp53(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp53(site) = Div(tmp52(site), tmp27(site), 999) - 1
        End If
    Next site

    ' 184.PutTestResult.HL_CS3RG_ZC1
    Call ResultAdd("HL_CS3RG_ZC1", tmp53)

' #### HL_CS3RG_ZC3 ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3RG_ZC3

    ' 1.Clamp.HL_CS3RG_ZC3

    ' 2.Median.HL_CS3RG_ZC3

    ' 3.Median.HL_CS3RG_ZC3

    ' 175.Average_FA.HL_CS3RG_ZC3

    ' 176.GetAverage_Color.HL_CS3RG_ZC3

    ' 177.GetAverage_Color.HL_CS3RG_ZC3

    ' 178.åvéZéÆï]âø.HL_CS3RG_ZC3

    ' 179.Average_FA.HL_CS3RG_ZC3
    Dim tmp54_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "HL_TEMP_Bayer2x4_ZONEC3", EEE_COLOR_ALL, tmp54_0)
    Dim tmp55 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp55, tmp54_0)

    ' 180.GetAverage_Color.HL_CS3RG_ZC3
    Dim tmp56(nSite) As Double
    Call GetAverage_Color(tmp56, tmp55, "Gr1", "Gb1", "Gr2", "Gb2")

    ' 181.GetAverage_Color.HL_CS3RG_ZC3
    Dim tmp57(nSite) As Double
    Call GetAverage_Color(tmp57, tmp55, "R1", "R2")

    ' 182.åvéZéÆï]âø.HL_CS3RG_ZC3
    Dim tmp58(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp58(site) = Div(tmp57(site), tmp56(site), 999)
        End If
    Next site

    ' 183.åvéZéÆï]âø.HL_CS3RG_ZC3
    Dim tmp59(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp59(site) = Div(tmp58(site), tmp27(site), 999) - 1
        End If
    Next site

    ' 184.PutTestResult.HL_CS3RG_ZC3
    Call ResultAdd("HL_CS3RG_ZC3", tmp59)

' #### HL_CS3RG_ZR1 ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3RG_ZR1

    ' 1.Clamp.HL_CS3RG_ZR1

    ' 2.Median.HL_CS3RG_ZR1

    ' 3.Median.HL_CS3RG_ZR1

    ' 175.Average_FA.HL_CS3RG_ZR1

    ' 176.GetAverage_Color.HL_CS3RG_ZR1

    ' 177.GetAverage_Color.HL_CS3RG_ZR1

    ' 178.åvéZéÆï]âø.HL_CS3RG_ZR1

    ' 179.Average_FA.HL_CS3RG_ZR1
    Dim tmp60_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "HL_TEMP_Bayer2x4_ZONER1", EEE_COLOR_ALL, tmp60_0)
    Dim tmp61 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp61, tmp60_0)

    ' 180.GetAverage_Color.HL_CS3RG_ZR1
    Dim tmp62(nSite) As Double
    Call GetAverage_Color(tmp62, tmp61, "Gr1", "Gb1", "Gr2", "Gb2")

    ' 181.GetAverage_Color.HL_CS3RG_ZR1
    Dim tmp63(nSite) As Double
    Call GetAverage_Color(tmp63, tmp61, "R1", "R2")

    ' 182.åvéZéÆï]âø.HL_CS3RG_ZR1
    Dim tmp64(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp64(site) = Div(tmp63(site), tmp62(site), 999)
        End If
    Next site

    ' 183.åvéZéÆï]âø.HL_CS3RG_ZR1
    Dim tmp65(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp65(site) = Div(tmp64(site), tmp27(site), 999) - 1
        End If
    Next site

    ' 184.PutTestResult.HL_CS3RG_ZR1
    Call ResultAdd("HL_CS3RG_ZR1", tmp65)

' #### HL_CS3RG_ZR2 ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3RG_ZR2

    ' 1.Clamp.HL_CS3RG_ZR2

    ' 2.Median.HL_CS3RG_ZR2

    ' 3.Median.HL_CS3RG_ZR2

    ' 175.Average_FA.HL_CS3RG_ZR2

    ' 176.GetAverage_Color.HL_CS3RG_ZR2

    ' 177.GetAverage_Color.HL_CS3RG_ZR2

    ' 178.åvéZéÆï]âø.HL_CS3RG_ZR2

    ' 179.Average_FA.HL_CS3RG_ZR2
    Dim tmp66_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "HL_TEMP_Bayer2x4_ZONER2", EEE_COLOR_ALL, tmp66_0)
    Dim tmp67 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp67, tmp66_0)

    ' 180.GetAverage_Color.HL_CS3RG_ZR2
    Dim tmp68(nSite) As Double
    Call GetAverage_Color(tmp68, tmp67, "Gr1", "Gb1", "Gr2", "Gb2")

    ' 181.GetAverage_Color.HL_CS3RG_ZR2
    Dim tmp69(nSite) As Double
    Call GetAverage_Color(tmp69, tmp67, "R1", "R2")

    ' 182.åvéZéÆï]âø.HL_CS3RG_ZR2
    Dim tmp70(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp70(site) = Div(tmp69(site), tmp68(site), 999)
        End If
    Next site

    ' 183.åvéZéÆï]âø.HL_CS3RG_ZR2
    Dim tmp71(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp71(site) = Div(tmp70(site), tmp27(site), 999) - 1
        End If
    Next site

    ' 184.PutTestResult.HL_CS3RG_ZR2
    Call ResultAdd("HL_CS3RG_ZR2", tmp71)

' #### HL_CS3RG_ZR3 ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3RG_ZR3

    ' 1.Clamp.HL_CS3RG_ZR3

    ' 2.Median.HL_CS3RG_ZR3

    ' 3.Median.HL_CS3RG_ZR3

    ' 175.Average_FA.HL_CS3RG_ZR3

    ' 176.GetAverage_Color.HL_CS3RG_ZR3

    ' 177.GetAverage_Color.HL_CS3RG_ZR3

    ' 178.åvéZéÆï]âø.HL_CS3RG_ZR3

    ' 179.Average_FA.HL_CS3RG_ZR3
    Dim tmp72_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "HL_TEMP_Bayer2x4_ZONER3", EEE_COLOR_ALL, tmp72_0)
        Call ReleasePlane(sPlane3)
    Dim tmp73 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp73, tmp72_0)

    ' 180.GetAverage_Color.HL_CS3RG_ZR3
    Dim tmp74(nSite) As Double
    Call GetAverage_Color(tmp74, tmp73, "Gr1", "Gb1", "Gr2", "Gb2")

    ' 181.GetAverage_Color.HL_CS3RG_ZR3
    Dim tmp75(nSite) As Double
    Call GetAverage_Color(tmp75, tmp73, "R1", "R2")

    ' 182.åvéZéÆï]âø.HL_CS3RG_ZR3
    Dim tmp76(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp76(site) = Div(tmp75(site), tmp74(site), 999)
        End If
    Next site

    ' 183.åvéZéÆï]âø.HL_CS3RG_ZR3
    Dim tmp77(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp77(site) = Div(tmp76(site), tmp27(site), 999) - 1
        End If
    Next site

    ' 184.PutTestResult.HL_CS3RG_ZR3
    Call ResultAdd("HL_CS3RG_ZR3", tmp77)

' #### HL_CS3BG_ZL1 ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3BG_ZL1

    ' 1.Clamp.HL_CS3BG_ZL1

    ' 2.Median.HL_CS3BG_ZL1

    ' 3.Median.HL_CS3BG_ZL1

    ' 175.Average_FA.HL_CS3BG_ZL1

    ' 176.GetAverage_Color.HL_CS3BG_ZL1

    ' 177.GetAverage_Color.HL_CS3BG_ZL1

    ' 178.åvéZéÆï]âø.HL_CS3BG_ZL1

    ' 179.Average_FA.HL_CS3BG_ZL1

    ' 180.GetAverage_Color.HL_CS3BG_ZL1

    ' 181.GetAverage_Color.HL_CS3BG_ZL1
    Dim tmp78(nSite) As Double
    Call GetAverage_Color(tmp78, tmp31, "B1", "B2")

    ' 182.åvéZéÆï]âø.HL_CS3BG_ZL1
    Dim tmp79(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp79(site) = Div(tmp78(site), tmp32(site), 999)
        End If
    Next site

    ' 183.åvéZéÆï]âø.HL_CS3BG_ZL1
    Dim tmp80(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp80(site) = Div(tmp79(site), tmp29(site), 999) - 1
        End If
    Next site

    ' 184.PutTestResult.HL_CS3BG_ZL1
    Call ResultAdd("HL_CS3BG_ZL1", tmp80)

' #### HL_CS3BG_ZL2 ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3BG_ZL2

    ' 1.Clamp.HL_CS3BG_ZL2

    ' 2.Median.HL_CS3BG_ZL2

    ' 3.Median.HL_CS3BG_ZL2

    ' 175.Average_FA.HL_CS3BG_ZL2

    ' 176.GetAverage_Color.HL_CS3BG_ZL2

    ' 177.GetAverage_Color.HL_CS3BG_ZL2

    ' 178.åvéZéÆï]âø.HL_CS3BG_ZL2

    ' 179.Average_FA.HL_CS3BG_ZL2

    ' 180.GetAverage_Color.HL_CS3BG_ZL2

    ' 181.GetAverage_Color.HL_CS3BG_ZL2
    Dim tmp81(nSite) As Double
    Call GetAverage_Color(tmp81, tmp37, "B1", "B2")

    ' 182.åvéZéÆï]âø.HL_CS3BG_ZL2
    Dim tmp82(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp82(site) = Div(tmp81(site), tmp38(site), 999)
        End If
    Next site

    ' 183.åvéZéÆï]âø.HL_CS3BG_ZL2
    Dim tmp83(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp83(site) = Div(tmp82(site), tmp29(site), 999) - 1
        End If
    Next site

    ' 184.PutTestResult.HL_CS3BG_ZL2
    Call ResultAdd("HL_CS3BG_ZL2", tmp83)

' #### HL_CS3BG_ZL3 ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3BG_ZL3

    ' 1.Clamp.HL_CS3BG_ZL3

    ' 2.Median.HL_CS3BG_ZL3

    ' 3.Median.HL_CS3BG_ZL3

    ' 175.Average_FA.HL_CS3BG_ZL3

    ' 176.GetAverage_Color.HL_CS3BG_ZL3

    ' 177.GetAverage_Color.HL_CS3BG_ZL3

    ' 178.åvéZéÆï]âø.HL_CS3BG_ZL3

    ' 179.Average_FA.HL_CS3BG_ZL3

    ' 180.GetAverage_Color.HL_CS3BG_ZL3

    ' 181.GetAverage_Color.HL_CS3BG_ZL3
    Dim tmp84(nSite) As Double
    Call GetAverage_Color(tmp84, tmp43, "B1", "B2")

    ' 182.åvéZéÆï]âø.HL_CS3BG_ZL3
    Dim tmp85(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp85(site) = Div(tmp84(site), tmp44(site), 999)
        End If
    Next site

    ' 183.åvéZéÆï]âø.HL_CS3BG_ZL3
    Dim tmp86(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp86(site) = Div(tmp85(site), tmp29(site), 999) - 1
        End If
    Next site

    ' 184.PutTestResult.HL_CS3BG_ZL3
    Call ResultAdd("HL_CS3BG_ZL3", tmp86)

' #### HL_CS3BG_ZC1 ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3BG_ZC1

    ' 1.Clamp.HL_CS3BG_ZC1

    ' 2.Median.HL_CS3BG_ZC1

    ' 3.Median.HL_CS3BG_ZC1

    ' 175.Average_FA.HL_CS3BG_ZC1

    ' 176.GetAverage_Color.HL_CS3BG_ZC1

    ' 177.GetAverage_Color.HL_CS3BG_ZC1

    ' 178.åvéZéÆï]âø.HL_CS3BG_ZC1

    ' 179.Average_FA.HL_CS3BG_ZC1

    ' 180.GetAverage_Color.HL_CS3BG_ZC1

    ' 181.GetAverage_Color.HL_CS3BG_ZC1
    Dim tmp87(nSite) As Double
    Call GetAverage_Color(tmp87, tmp49, "B1", "B2")

    ' 182.åvéZéÆï]âø.HL_CS3BG_ZC1
    Dim tmp88(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp88(site) = Div(tmp87(site), tmp50(site), 999)
        End If
    Next site

    ' 183.åvéZéÆï]âø.HL_CS3BG_ZC1
    Dim tmp89(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp89(site) = Div(tmp88(site), tmp29(site), 999) - 1
        End If
    Next site

    ' 184.PutTestResult.HL_CS3BG_ZC1
    Call ResultAdd("HL_CS3BG_ZC1", tmp89)

' #### HL_CS3BG_ZC3 ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3BG_ZC3

    ' 1.Clamp.HL_CS3BG_ZC3

    ' 2.Median.HL_CS3BG_ZC3

    ' 3.Median.HL_CS3BG_ZC3

    ' 175.Average_FA.HL_CS3BG_ZC3

    ' 176.GetAverage_Color.HL_CS3BG_ZC3

    ' 177.GetAverage_Color.HL_CS3BG_ZC3

    ' 178.åvéZéÆï]âø.HL_CS3BG_ZC3

    ' 179.Average_FA.HL_CS3BG_ZC3

    ' 180.GetAverage_Color.HL_CS3BG_ZC3

    ' 181.GetAverage_Color.HL_CS3BG_ZC3
    Dim tmp90(nSite) As Double
    Call GetAverage_Color(tmp90, tmp55, "B1", "B2")

    ' 182.åvéZéÆï]âø.HL_CS3BG_ZC3
    Dim tmp91(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp91(site) = Div(tmp90(site), tmp56(site), 999)
        End If
    Next site

    ' 183.åvéZéÆï]âø.HL_CS3BG_ZC3
    Dim tmp92(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp92(site) = Div(tmp91(site), tmp29(site), 999) - 1
        End If
    Next site

    ' 184.PutTestResult.HL_CS3BG_ZC3
    Call ResultAdd("HL_CS3BG_ZC3", tmp92)

' #### HL_CS3BG_ZR1 ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3BG_ZR1

    ' 1.Clamp.HL_CS3BG_ZR1

    ' 2.Median.HL_CS3BG_ZR1

    ' 3.Median.HL_CS3BG_ZR1

    ' 175.Average_FA.HL_CS3BG_ZR1

    ' 176.GetAverage_Color.HL_CS3BG_ZR1

    ' 177.GetAverage_Color.HL_CS3BG_ZR1

    ' 178.åvéZéÆï]âø.HL_CS3BG_ZR1

    ' 179.Average_FA.HL_CS3BG_ZR1

    ' 180.GetAverage_Color.HL_CS3BG_ZR1

    ' 181.GetAverage_Color.HL_CS3BG_ZR1
    Dim tmp93(nSite) As Double
    Call GetAverage_Color(tmp93, tmp61, "B1", "B2")

    ' 182.åvéZéÆï]âø.HL_CS3BG_ZR1
    Dim tmp94(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp94(site) = Div(tmp93(site), tmp62(site), 999)
        End If
    Next site

    ' 183.åvéZéÆï]âø.HL_CS3BG_ZR1
    Dim tmp95(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp95(site) = Div(tmp94(site), tmp29(site), 999) - 1
        End If
    Next site

    ' 184.PutTestResult.HL_CS3BG_ZR1
    Call ResultAdd("HL_CS3BG_ZR1", tmp95)

' #### HL_CS3BG_ZR2 ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3BG_ZR2

    ' 1.Clamp.HL_CS3BG_ZR2

    ' 2.Median.HL_CS3BG_ZR2

    ' 3.Median.HL_CS3BG_ZR2

    ' 175.Average_FA.HL_CS3BG_ZR2

    ' 176.GetAverage_Color.HL_CS3BG_ZR2

    ' 177.GetAverage_Color.HL_CS3BG_ZR2

    ' 178.åvéZéÆï]âø.HL_CS3BG_ZR2

    ' 179.Average_FA.HL_CS3BG_ZR2

    ' 180.GetAverage_Color.HL_CS3BG_ZR2

    ' 181.GetAverage_Color.HL_CS3BG_ZR2
    Dim tmp96(nSite) As Double
    Call GetAverage_Color(tmp96, tmp67, "B1", "B2")

    ' 182.åvéZéÆï]âø.HL_CS3BG_ZR2
    Dim tmp97(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp97(site) = Div(tmp96(site), tmp68(site), 999)
        End If
    Next site

    ' 183.åvéZéÆï]âø.HL_CS3BG_ZR2
    Dim tmp98(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp98(site) = Div(tmp97(site), tmp29(site), 999) - 1
        End If
    Next site

    ' 184.PutTestResult.HL_CS3BG_ZR2
    Call ResultAdd("HL_CS3BG_ZR2", tmp98)

' #### HL_CS3BG_ZR3 ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_CS3BG_ZR3

    ' 1.Clamp.HL_CS3BG_ZR3

    ' 2.Median.HL_CS3BG_ZR3

    ' 3.Median.HL_CS3BG_ZR3

    ' 175.Average_FA.HL_CS3BG_ZR3

    ' 176.GetAverage_Color.HL_CS3BG_ZR3

    ' 177.GetAverage_Color.HL_CS3BG_ZR3

    ' 178.åvéZéÆï]âø.HL_CS3BG_ZR3

    ' 179.Average_FA.HL_CS3BG_ZR3

    ' 180.GetAverage_Color.HL_CS3BG_ZR3

    ' 181.GetAverage_Color.HL_CS3BG_ZR3
    Dim tmp99(nSite) As Double
    Call GetAverage_Color(tmp99, tmp73, "B1", "B2")

    ' 182.åvéZéÆï]âø.HL_CS3BG_ZR3
    Dim tmp100(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp100(site) = Div(tmp99(site), tmp74(site), 999)
        End If
    Next site

    ' 183.åvéZéÆï]âø.HL_CS3BG_ZR3
    Dim tmp101(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp101(site) = Div(tmp100(site), tmp29(site), 999) - 1
        End If
    Next site

    ' 184.PutTestResult.HL_CS3BG_ZR3
    Call ResultAdd("HL_CS3BG_ZR3", tmp101)

' #### HL_ABMAXL ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_ABMAXL

    ' 1.Clamp.HL_ABMAXL

    ' 2.Lab_RGB_Separate.HL_ABMAXL
    
    Call labProc_Initialize("HL_TEMP_Bayer2x4", "HL_TEMP_Bayer2x4_FULL", "HL_TEMP_Bayer2x4_ZONE3", "HL_TEMP_Bayer2x4_VOPB")
    Const MYPLANE_BAYER As String = "pBayer"
    Const MYPLANE_BAYER_R As String = "pBayerR"
    Const MYPLANE_BAYER_G As String = "pBayerG"
    Const MYPLANE_BAYER_B As String = "pBayerB"
    Dim compFactor As Long
    compFactor = labProc_ReturnCompFactor
    Dim bayerPlane As CImgPlane
    Call GetFreePlane(bayerPlane, "pBayer", idpDepthS16, , "bayerPlane")
    Call Copy(sPlane1, "HL_TEMP_Bayer2x4_ZONE3", EEE_COLOR_FLAT, bayerPlane, "BAYER_FULL", EEE_COLOR_FLAT)
        Call ReleasePlane(sPlane1)
    Dim bayerPlaneR As CImgPlane
    Call GetFreePlane(bayerPlaneR, "pBayerR", idpDepthS16, , "R extracted plane")
    Call Copy(bayerPlane, "BAYER_FULL", "R", bayerPlaneR, "BAYER_R_FULL", "R")
    Dim bayerPlaneG As CImgPlane
    Call GetFreePlane(bayerPlaneG, "pBayerG", idpDepthS16, , "G extracted plane")
    Call MultiMean(bayerPlane, "BAYER_FULL", "GR", bayerPlaneG, "BAYER_G_FULL", "GR", idpMultiMeanFuncMean, 1, 2)
    Dim bayerPlaneB As CImgPlane
    Call GetFreePlane(bayerPlaneB, "pBayerB", idpDepthS16, , "B extracted plane")
    Call Copy(bayerPlane, "BAYER_FULL", "B", bayerPlaneB, "BAYER_B_FULL", "B")
        Call ReleasePlane(bayerPlane)

    ' 3.Median(ägí£Ç»Çµ).HL_ABMAXL
    Dim sPlane9 As CImgPlane
    Call GetFreePlane(sPlane9, "pBayerR", idpDepthS16, , "sPlane9")
    Call Median(bayerPlaneR, "BAYER_R_FULL", EEE_COLOR_FLAT, sPlane9, "BAYER_R_FULL", EEE_COLOR_FLAT, 5, 1)
        Call ReleasePlane(bayerPlaneR)

    ' 4.Median(ägí£Ç»Çµ).HL_ABMAXL
    Dim sPlane10 As CImgPlane
    Call GetFreePlane(sPlane10, "pBayerR", idpDepthS16, , "sPlane10")
    Call Median(sPlane9, "BAYER_R_FULL", EEE_COLOR_FLAT, sPlane10, "BAYER_R_FULL", EEE_COLOR_FLAT, 1, 5)
        Call ReleasePlane(sPlane9)

    ' 5.Median(ägí£Ç»Çµ).HL_ABMAXL
    Dim sPlane11 As CImgPlane
    Call GetFreePlane(sPlane11, "pBayerG", idpDepthS16, , "sPlane11")
    Call Median(bayerPlaneG, "BAYER_G_FULL", EEE_COLOR_FLAT, sPlane11, "BAYER_G_FULL", EEE_COLOR_FLAT, 5, 1)
        Call ReleasePlane(bayerPlaneG)

    ' 6.Median(ägí£Ç»Çµ).HL_ABMAXL
    Dim sPlane12 As CImgPlane
    Call GetFreePlane(sPlane12, "pBayerG", idpDepthS16, , "sPlane12")
    Call Median(sPlane11, "BAYER_G_FULL", EEE_COLOR_FLAT, sPlane12, "BAYER_G_FULL", EEE_COLOR_FLAT, 1, 5)
        Call ReleasePlane(sPlane11)

    ' 7.Median(ägí£Ç»Çµ).HL_ABMAXL
    Dim sPlane13 As CImgPlane
    Call GetFreePlane(sPlane13, "pBayerB", idpDepthS16, , "sPlane13")
    Call Median(bayerPlaneB, "BAYER_B_FULL", EEE_COLOR_FLAT, sPlane13, "BAYER_B_FULL", EEE_COLOR_FLAT, 5, 1)
        Call ReleasePlane(bayerPlaneB)

    ' 8.Median(ägí£Ç»Çµ).HL_ABMAXL
    Dim sPlane14 As CImgPlane
    Call GetFreePlane(sPlane14, "pBayerB", idpDepthS16, , "sPlane14")
    Call Median(sPlane13, "BAYER_B_FULL", EEE_COLOR_FLAT, sPlane14, "BAYER_B_FULL", EEE_COLOR_FLAT, 1, 5)
        Call ReleasePlane(sPlane13)

    ' 9.Lab_RGB_Local.HL_ABMAXL
    Const MYPLANE_LOCAL As String = "pLabL"
    Const MYPLANE_COMP_BAYER_G As String = "pcBayerG"
    Const MYPLANE_COMP_BAYER_B As String = "pcBayerB"
    Dim rISrcPlane As CImgPlane
    Call GetFreePlane(rISrcPlane, "pLabL", idpDepthS16, , "R (int) Source Plane")
    Dim gISrcPlane As CImgPlane
    Call GetFreePlane(gISrcPlane, "pLabL", idpDepthS16, , "G (int) Source Plane")
    Dim bISrcPlane As CImgPlane
    Call GetFreePlane(bISrcPlane, "pLabL", idpDepthS16, , "B (int) Source Plane")
    Dim gCompPlane As CImgPlane
    Call GetFreePlane(gCompPlane, "pcBayerG", idpDepthS16, , "G Compressed Plane")
    Dim bCompPlane As CImgPlane
    Call GetFreePlane(bCompPlane, "pcBayerB", idpDepthS16, , "B Compressed Plane")
    Call MultiMean(sPlane10, "BAYER_R_ZONE2D", EEE_COLOR_FLAT, rISrcPlane, "LABZ2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, COMP_PRIMARY_TO_LOCAL, COMP_PRIMARY_TO_LOCAL)
        Call ReleasePlane(sPlane10)
    Call MultiMean(sPlane12, "BAYER_G_ZONE2D", EEE_COLOR_FLAT, gCompPlane, "PCBAYER_G_Z2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, COMP_PRIMARY_TO_LOCAL, COMP_PRIMARY_TO_LOCAL)
        Call ReleasePlane(sPlane12)
    Call Copy(gCompPlane, "PCBAYER_G_Z2D", EEE_COLOR_FLAT, gISrcPlane, "LABZ2D", EEE_COLOR_FLAT)
        Call ReleasePlane(gCompPlane)
    Call MultiMean(sPlane14, "BAYER_B_ZONE2D", EEE_COLOR_FLAT, bCompPlane, "PCBAYER_B_Z2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, COMP_PRIMARY_TO_LOCAL, COMP_PRIMARY_TO_LOCAL)
        Call ReleasePlane(sPlane14)
    Call Copy(bCompPlane, "PCBAYER_B_Z2D", EEE_COLOR_FLAT, bISrcPlane, "LABZ2D", EEE_COLOR_FLAT)
        Call ReleasePlane(bCompPlane)

    ' 10.FlatFielding.HL_ABMAXL
    Dim rFSrcPlane As CImgPlane
     Call GetFreePlane(rFSrcPlane, "pLabL", idpDepthF32, , "R (Float) Source Plane")
     Call Copy(rISrcPlane, "LABZ2D", EEE_COLOR_FLAT, rFSrcPlane, "LABZ2D", EEE_COLOR_FLAT)
        Call ReleasePlane(rISrcPlane)
     Dim gFSrcPlane As CImgPlane
     Call GetFreePlane(gFSrcPlane, "pLabL", idpDepthF32, , "G (Float) Source Plane")
     Call Copy(gISrcPlane, "LABZ2D", EEE_COLOR_FLAT, gFSrcPlane, "LABZ2D", EEE_COLOR_FLAT)
        Call ReleasePlane(gISrcPlane)
     Dim bFSrcPlane As CImgPlane
     Call GetFreePlane(bFSrcPlane, "pLabL", idpDepthF32, , "B (Float) Source Plane")
     Call Copy(bISrcPlane, "LABZ2D", EEE_COLOR_FLAT, bFSrcPlane, "LABZ2D", EEE_COLOR_FLAT)
        Call ReleasePlane(bISrcPlane)
     Dim ffrPlane As CImgPlane
     Dim ffgPlane As CImgPlane
     Dim ffbPlane As CImgPlane
     Call procLab_GetFlatFields(ffrPlane, ffgPlane, ffbPlane)
     Dim rNoShdPlane As CImgPlane
     Call GetFreePlane(rNoShdPlane, "pLabL", idpDepthF32, , "R Shading Removed")
     Call Divide(rFSrcPlane, "LABZ2D", EEE_COLOR_FLAT, ffrPlane, "LABZ2D", EEE_COLOR_FLAT, rNoShdPlane, "LABZ2D", EEE_COLOR_FLAT)
        Call ReleasePlane(rFSrcPlane)
     Dim gNoShdPlane As CImgPlane
     Call GetFreePlane(gNoShdPlane, "pLabL", idpDepthF32, , "G Shading Removed")
     Call Divide(gFSrcPlane, "LABZ2D", EEE_COLOR_FLAT, ffgPlane, "LABZ2D", EEE_COLOR_FLAT, gNoShdPlane, "LABZ2D", EEE_COLOR_FLAT)
        Call ReleasePlane(gFSrcPlane)
     Dim bNoShdPlane As CImgPlane
     Call GetFreePlane(bNoShdPlane, "pLabL", idpDepthF32, , "B Shading Removed")
     Call Divide(bFSrcPlane, "LABZ2D", EEE_COLOR_FLAT, ffbPlane, "LABZ2D", EEE_COLOR_FLAT, bNoShdPlane, "LABZ2D", EEE_COLOR_FLAT)
        Call ReleasePlane(bFSrcPlane)

    ' 11.LowPassFilter.HL_ABMAXL
    Dim LfPlane As CImgPlane
    Dim AfPlane As CImgPlane
    Dim BfPlane As CImgPlane
    Call GetFreePlane(LfPlane, "pLabL", idpDepthF32, , "L Plane")
    Call GetFreePlane(AfPlane, "pLabL", idpDepthF32, False, "a Plane")
    Call GetFreePlane(BfPlane, "pLabL", idpDepthF32, False, "b Plane")
    Call labProc_RGB2LabDirect(rNoShdPlane, gNoShdPlane, bNoShdPlane, LfPlane, AfPlane, BfPlane, "LABZ2D", "LABZ0", 1)
        Call ReleasePlane(rNoShdPlane)
        Call ReleasePlane(gNoShdPlane)
        Call ReleasePlane(bNoShdPlane)
    Call labProc_ApplyLPF(LfPlane, "LABZ2D", "kernel_LowPassH_ColorFloat", "kernel_LowPassV_ColorFloat")
    Call labProc_ApplyLPF(AfPlane, "LABZ2D", "kernel_LowPassH_ColorFloat", "kernel_LowPassV_ColorFloat")
    Call labProc_ApplyLPF(BfPlane, "LABZ2D", "kernel_LowPassH_ColorFloat", "kernel_LowPassV_ColorFloat")

    ' 12.êFÉÄÉâ.HL_ABMAXL
    Dim returnLocalRowMaxC(nSite) As Double
    Dim returnLocalColMaxC(nSite) As Double
    Call labProc_abmaxRow(AfPlane, BfPlane, LOCAL_DIFF_SIZE, "LABZ2D", "LABZ2D_LOCAL_JUDGE_ROW", returnLocalRowMaxC)
    Call labProc_abmaxCol(AfPlane, BfPlane, "LABZ2D_LOCAL_SOURCE_COL", "LABZ2D_LOCAL_TARGET_COL", "LABZ2D_LOCAL_JUDGE_COL", returnLocalColMaxC)
        Call ReleasePlane(AfPlane)
        Call ReleasePlane(BfPlane)
    Dim tmp102(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            If returnLocalRowMaxC(site) > returnLocalColMaxC(site) Then
                tmp102(site) = returnLocalRowMaxC(site)
            Else
                tmp102(site) = returnLocalColMaxC(site)
            End If
        End If
    Next site

    ' 13.PutTestResult.HL_ABMAXL
    Call ResultAdd("HL_ABMAXL", tmp102)

' #### HL_LMAXL ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_LMAXL

    ' 1.Clamp.HL_LMAXL

    ' 2.Lab_RGB_Separate.HL_LMAXL
    
    ' 3.Median(ägí£Ç»Çµ).HL_LMAXL

    ' 4.Median(ägí£Ç»Çµ).HL_LMAXL

    ' 5.Median(ägí£Ç»Çµ).HL_LMAXL

    ' 6.Median(ägí£Ç»Çµ).HL_LMAXL

    ' 7.Median(ägí£Ç»Çµ).HL_LMAXL

    ' 8.Median(ägí£Ç»Çµ).HL_LMAXL

    ' 9.Lab_RGB_Local.HL_LMAXL

    ' 10.FlatFielding.HL_LMAXL

    ' 11.LowPassFilter.HL_LMAXL

    ' 16.Lab_ç≈ëÂíl.HL_LMAXL
    Const PROCESS_ZONE As String = "LABZ2D"
    Const HDIF_ZONE_SOURCE As String = "LABZ2D_LOCAL_SOURCE_COL"
    Const HDIF_ZONE_TARGET As String = "LABZ2D_LOCAL_TARGET_COL"
    Const HDIF_LMAXL_JUDGE As String = "LABZ2D_LOCAL_JUDGE_COL"
    Const VDIF_LMAXL_JUDGE As String = "LABZ2D_LOCAL_JUDGE_ROW"
    Dim difPlane As CImgPlane
    Call GetFreePlane(difPlane, LfPlane.planeGroup, idpDepthF32, False, "Diff Plane")
    Call SubRows(LfPlane, PROCESS_ZONE, EEE_COLOR_FLAT, difPlane, PROCESS_ZONE, EEE_COLOR_FLAT, LOCAL_DIFF_SIZE)
    Dim tmp103(nSite) As Double
    Call AbsMax(difPlane, VDIF_LMAXL_JUDGE, EEE_COLOR_FLAT, tmp103)
    Dim workPlane As CImgPlane
    Call GetFreePlane(workPlane, LfPlane.planeGroup, idpDepthF32, False, "Work Plane")
    Call Copy(LfPlane, PROCESS_ZONE, EEE_COLOR_FLAT, workPlane, PROCESS_ZONE, EEE_COLOR_FLAT)
    Call Subtract(LfPlane, HDIF_ZONE_SOURCE, EEE_COLOR_FLAT, workPlane, HDIF_ZONE_TARGET, EEE_COLOR_FLAT, difPlane, HDIF_ZONE_SOURCE, EEE_COLOR_FLAT)
        Call ReleasePlane(workPlane)
    Dim tmpreturnLmaxl(nSite) As Double
    Call AbsMax(difPlane, HDIF_LMAXL_JUDGE, EEE_COLOR_FLAT, tmpreturnLmaxl)
        Call ReleasePlane(difPlane)
    Dim tmp104(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            If Abs(tmp103(site)) < Abs(tmpreturnLmaxl(site)) Then
                tmp104(site) = Abs(tmpreturnLmaxl(site))
            Else
                tmp104(site) = Abs(tmp103(site))
            End If
        End If
    Next site

    ' 17.PutTestResult.HL_LMAXL
    Call ResultAdd("HL_LMAXL", tmp104)

' #### HL_LDENL ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL_LDENL

    ' 1.Clamp.HL_LDENL

    ' 2.Lab_RGB_Separate.HL_LDENL
    
    ' 3.Median(ägí£Ç»Çµ).HL_LDENL

    ' 4.Median(ägí£Ç»Çµ).HL_LDENL

    ' 5.Median(ägí£Ç»Çµ).HL_LDENL

    ' 6.Median(ägí£Ç»Çµ).HL_LDENL

    ' 7.Median(ägí£Ç»Çµ).HL_LDENL

    ' 8.Median(ägí£Ç»Çµ).HL_LDENL

    ' 9.Lab_RGB_Local.HL_LDENL

    ' 10.FlatFielding.HL_LDENL

    ' 11.LowPassFilter.HL_LDENL

    ' 14.ÉnÉLÉÄÉâ.HL_LDENL
    Dim tmp105(nSite) As Double
    Call Lab_Haki(LfPlane, 1.2, tmp105)
        Call ReleasePlane(LfPlane)

    ' 15.PutTestResult.HL_LDENL
    Call ResultAdd("HL_LDENL", tmp105)

' #### TMP_LV_EBD ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.TMP_LV_EBD

    ' 3.Average_FA.TMP_LV_EBD
    Dim tmp106_0 As CImgColorAllResult
    Call Average_FA(HL_ERR_Plane, "HL_TEMP_Bayer2x4_ZONE_EBD", EEE_COLOR_ALL, tmp106_0)
    Dim tmp107 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp107, tmp106_0)

    ' 4.GetAverage_Color.TMP_LV_EBD
    Dim tmp108(nSite) As Double
    Call GetAverage_Color(tmp108, tmp107, "-")

    ' 6.åvéZéÆï]âø.TMP_LV_EBD
    Dim tmp109(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp109(site) = 4
        End If
    Next site

    ' 7.åvéZéÆï]âø.TMP_LV_EBD
    Dim tmp110(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp110(site) = Div(tmp108(site), tmp109(site), 999)
        End If
    Next site

    ' 8.PutTestResult.TMP_LV_EBD
    Call ResultAdd("TMP_LV_EBD", tmp110)

' #### TMP_LV ####

    ' 9.åvéZéÆï]âø.TMP_LV
    Dim tmp111(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp111(site) = 1.527559
        End If
    Next site

    ' 10.åvéZéÆï]âø.TMP_LV
    Dim tmp112(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp112(site) = 541
        End If
    Next site

    ' 11.åvéZéÆï]âø.TMP_LV
    Dim tmp113(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp113(site) = tmp110(site) * tmp111(site) + tmp112(site)
        End If
    Next site

    ' 12.PutTestResult.TMP_LV
    Call ResultAdd("TMP_LV", tmp113)

' #### TMP_SFAV ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.TMP_SFAV

    ' 3.Average_FA.TMP_SFAV
    Dim tmp114_0 As CImgColorAllResult
    Call Average_FA(HL_ERR_Plane, "HL_TEMP_Bayer2x4_ZONE_SFG", EEE_COLOR_ALL, tmp114_0)
    Dim tmp115 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp115, tmp114_0)

    ' 4.GetAverage_Color.TMP_SFAV
    Dim tmp116(nSite) As Double
    Call GetAverage_Color(tmp116, tmp115, "-")

    ' 5.PutTestResult.TMP_SFAV
    Call ResultAdd("TMP_SFAV", tmp116)

' #### TMP_SFDF ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.TMP_SFDF

    ' 3.Average_FA.TMP_SFDF

    ' 4.GetAverage_Color.TMP_SFDF

    ' 13.åvéZéÆï]âø.TMP_SFDF
    Dim tmp117(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp117(site) = 64
        End If
    Next site

    ' 14.åvéZéÆï]âø.TMP_SFDF
    Dim tmp118(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp118(site) = 0.0003448275862
        End If
    Next site

    ' 15.åvéZéÆï]âø.TMP_SFDF
    Dim tmp119(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp119(site) = (tmp116(site) - tmp117(site)) * tmp118(site)
        End If
    Next site

    ' 16.PutTestResult.TMP_SFDF
    Call ResultAdd("TMP_SFDF", tmp119)

' #### TMP_SFSGA ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.TMP_SFSGA

    ' 3.Average_FA.TMP_SFSGA

    ' 4.GetAverage_Color.TMP_SFSGA

    ' 13.åvéZéÆï]âø.TMP_SFSGA

    ' 14.åvéZéÆï]âø.TMP_SFSGA

    ' 15.åvéZéÆï]âø.TMP_SFSGA

    ' 18.åvéZéÆï]âø.TMP_SFSGA
    Dim tmp120(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp120(site) = 0.278
        End If
    Next site

    ' 20.åvéZéÆï]âø.TMP_SFSGA
    Dim tmp121(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp121(site) = Div(tmp119(site), tmp120(site), 999)
        End If
    Next site

    ' 21.PutTestResult.TMP_SFSGA
    Call ResultAdd("TMP_SFSGA", tmp121)

' #### TMP_SLP ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.TMP_SLP

    ' 3.Average_FA.TMP_SLP

    ' 4.GetAverage_Color.TMP_SLP

    ' 13.åvéZéÆï]âø.TMP_SLP

    ' 14.åvéZéÆï]âø.TMP_SLP

    ' 15.åvéZéÆï]âø.TMP_SLP

    ' 18.åvéZéÆï]âø.TMP_SLP

    ' 20.åvéZéÆï]âø.TMP_SLP

    ' 22.åvéZéÆï]âø.TMP_SLP
    Dim tmp122(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp122(site) = 228.24
        End If
    Next site

    ' 24.åvéZéÆï]âø.TMP_SLP
    Dim tmp123(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp123(site) = tmp121(site) * tmp122(site)
        End If
    Next site

    ' 25.PutTestResult.TMP_SLP
    Call ResultAdd("TMP_SLP", tmp123)

' #### TMP_OFS ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.TMP_OFS

    ' 1.Average_FA.TMP_OFS

    ' 2.GetAverage_Color.TMP_OFS

    ' 3.Average_FA.TMP_OFS

    ' 4.GetAverage_Color.TMP_OFS

    ' 6.åvéZéÆï]âø.TMP_OFS

    ' 7.åvéZéÆï]âø.TMP_OFS

    ' 9.åvéZéÆï]âø.TMP_OFS

    ' 10.åvéZéÆï]âø.TMP_OFS

    ' 11.åvéZéÆï]âø.TMP_OFS

    ' 13.åvéZéÆï]âø.TMP_OFS

    ' 14.åvéZéÆï]âø.TMP_OFS

    ' 17.åvéZéÆï]âø.TMP_OFS

    ' 18.åvéZéÆï]âø.TMP_OFS

    ' 19.åvéZéÆï]âø.TMP_OFS

    ' 22.åvéZéÆï]âø.TMP_OFS

    ' 23.åvéZéÆï]âø.TMP_OFS

    ' 26.åvéZéÆï]âø.TMP_OFS
    Dim tmp124(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp124(site) = 0.6693
        End If
    Next site

    ' 27.åvéZéÆï]âø.TMP_OFS
    Dim tmp125(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp125(site) = tmp113(site) - tmp123(site) * tmp124(site)
        End If
    Next site

    ' 28.PutTestResult.TMP_OFS
    Call ResultAdd("TMP_OFS", tmp125)

End Function



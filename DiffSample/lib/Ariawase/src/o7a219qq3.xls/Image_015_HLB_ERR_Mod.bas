Attribute VB_Name = "Image_015_HLB_ERR_Mod"

Option Explicit

Public Function HLB_ERR_Process()

        Call PutImageInto_Common

' #### HLB_SENR ####

    Dim site As Long

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HLB_SENR
    Dim HLB_ERR_Param As CParamPlane
    Dim HLB_ERR_DevInfo As CDeviceConfigInfo
    Dim HLB_ERR_Plane As CImgPlane
    Set HLB_ERR_Param = TheParameterBank.Item("HLBImageTest_Acq1")
    Set HLB_ERR_DevInfo = TheDeviceProfiler.ConfigInfo("HLBImageTest_Acq1")
        Call TheParameterBank.Delete("HLBImageTest_Acq1")
    Set HLB_ERR_Plane = HLB_ERR_Param.plane

    ' 1.Clamp.HLB_SENR
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call Clamp(HLB_ERR_Plane, sPlane1, "Bayer2x4_VOPB")

    ' 2.Median.HLB_SENR
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE2D", 1, 5)
        Call ReleasePlane(sPlane1)

    ' 3.Median.HLB_SENR
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call MedianEx(sPlane2, sPlane3, "Bayer2x4_ZONE2D", 5, 1)
        Call ReleasePlane(sPlane2)

    ' 82.Average_FA.HLB_SENR
    Dim tmp1_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONE0", EEE_COLOR_ALL, tmp1_0)
    Dim tmp2 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp2, tmp1_0)

    ' 83.GetAverage_Color.HLB_SENR
    Dim tmp3(nSite) As Double
    Call GetAverage_Color(tmp3, tmp2, "R1", "R2")

    ' 233.LSBíËã`.HLB_SENR
    Dim HLB_ERR_LSB() As Double
     HLB_ERR_LSB = HLB_ERR_DevInfo.Lsb.AsDouble

    ' 238.LSBä∑éZ.HLB_SENR
    Dim tmp4(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp4(site) = tmp3(site) * HLB_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.HLB_SENR
    Call ResultAdd("HLB_SENR", tmp4)

' #### HLB_SENGR ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HLB_SENGR

    ' 1.Clamp.HLB_SENGR

    ' 2.Median.HLB_SENGR

    ' 3.Median.HLB_SENGR

    ' 82.Average_FA.HLB_SENGR

    ' 83.GetAverage_Color.HLB_SENGR
    Dim tmp5(nSite) As Double
    Call GetAverage_Color(tmp5, tmp2, "Gr1", "Gr2")

    ' 233.LSBíËã`.HLB_SENGR

    ' 238.LSBä∑éZ.HLB_SENGR
    Dim tmp6(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp6(site) = tmp5(site) * HLB_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.HLB_SENGR
    Call ResultAdd("HLB_SENGR", tmp6)

' #### HLB_SENGB ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HLB_SENGB

    ' 1.Clamp.HLB_SENGB

    ' 2.Median.HLB_SENGB

    ' 3.Median.HLB_SENGB

    ' 82.Average_FA.HLB_SENGB

    ' 83.GetAverage_Color.HLB_SENGB
    Dim tmp7(nSite) As Double
    Call GetAverage_Color(tmp7, tmp2, "Gb1", "Gb2")

    ' 233.LSBíËã`.HLB_SENGB

    ' 238.LSBä∑éZ.HLB_SENGB
    Dim tmp8(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp8(site) = tmp7(site) * HLB_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.HLB_SENGB
    Call ResultAdd("HLB_SENGB", tmp8)

' #### HLB_SENB ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HLB_SENB

    ' 1.Clamp.HLB_SENB

    ' 2.Median.HLB_SENB

    ' 3.Median.HLB_SENB

    ' 82.Average_FA.HLB_SENB

    ' 83.GetAverage_Color.HLB_SENB
    Dim tmp9(nSite) As Double
    Call GetAverage_Color(tmp9, tmp2, "B1", "B2")

    ' 233.LSBíËã`.HLB_SENB

    ' 238.LSBä∑éZ.HLB_SENB
    Dim tmp10(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp10(site) = tmp9(site) * HLB_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.HLB_SENB
    Call ResultAdd("HLB_SENB", tmp10)

' #### HLB_LC_Z22 ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HLB_LC_Z22

    ' 1.Clamp.HLB_LC_Z22

    ' 2.Median.HLB_LC_Z22

    ' 3.Median.HLB_LC_Z22

    ' 50.Average_FA.HLB_LC_Z22
    Dim tmp11_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONE22", EEE_COLOR_ALL, tmp11_0)
    Dim tmp12 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp12, tmp11_0)

    ' 51.GetAverage_Color.HLB_LC_Z22
    Dim tmp13(nSite) As Double
    Call GetAverage_Color(tmp13, tmp12, "Gr1", "Gr2")

    ' 52.GetAverage_Color.HLB_LC_Z22
    Dim tmp14(nSite) As Double
    Call GetAverage_Color(tmp14, tmp12, "Gb1", "Gb2")

    ' 53.åvéZéÆï]âø.HLB_LC_Z22
    Dim tmp15(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp15(site) = tmp13(site) - tmp14(site)
        End If
    Next site

    ' 54.åvéZéÆï]âø.HLB_LC_Z22
    Dim tmp16(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp16(site) = (tmp13(site) + tmp14(site)) / 2
        End If
    Next site

    ' 55.åvéZéÆï]âø.HLB_LC_Z22
    Dim tmp17(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp17(site) = Div(tmp15(site), tmp16(site), 999)
        End If
    Next site

    ' 56.PutTestResult.HLB_LC_Z22
    Call ResultAdd("HLB_LC_Z22", tmp17)

' #### HLB_LC_Z2D ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HLB_LC_Z2D

    ' 1.Clamp.HLB_LC_Z2D

    ' 2.Median.HLB_LC_Z2D

    ' 3.Median.HLB_LC_Z2D

    ' 50.Average_FA.HLB_LC_Z2D
    Dim tmp18_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, tmp18_0)
    Dim tmp19 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp19, tmp18_0)

    ' 51.GetAverage_Color.HLB_LC_Z2D
    Dim tmp20(nSite) As Double
    Call GetAverage_Color(tmp20, tmp19, "Gr1", "Gr2")

    ' 52.GetAverage_Color.HLB_LC_Z2D
    Dim tmp21(nSite) As Double
    Call GetAverage_Color(tmp21, tmp19, "Gb1", "Gb2")

    ' 53.åvéZéÆï]âø.HLB_LC_Z2D
    Dim tmp22(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp22(site) = tmp20(site) - tmp21(site)
        End If
    Next site

    ' 54.åvéZéÆï]âø.HLB_LC_Z2D
    Dim tmp23(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp23(site) = (tmp20(site) + tmp21(site)) / 2
        End If
    Next site

    ' 55.åvéZéÆï]âø.HLB_LC_Z2D
    Dim tmp24(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp24(site) = Div(tmp22(site), tmp23(site), 999)
        End If
    Next site

    ' 56.PutTestResult.HLB_LC_Z2D
    Call ResultAdd("HLB_LC_Z2D", tmp24)

' #### HLB_LC_ZLT ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HLB_LC_ZLT

    ' 1.Clamp.HLB_LC_ZLT

    ' 2.Median.HLB_LC_ZLT

    ' 3.Median.HLB_LC_ZLT

    ' 50.Average_FA.HLB_LC_ZLT
    Dim tmp25_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONELT", EEE_COLOR_ALL, tmp25_0)
    Dim tmp26 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp26, tmp25_0)

    ' 51.GetAverage_Color.HLB_LC_ZLT
    Dim tmp27(nSite) As Double
    Call GetAverage_Color(tmp27, tmp26, "Gr1", "Gr2")

    ' 52.GetAverage_Color.HLB_LC_ZLT
    Dim tmp28(nSite) As Double
    Call GetAverage_Color(tmp28, tmp26, "Gb1", "Gb2")

    ' 53.åvéZéÆï]âø.HLB_LC_ZLT
    Dim tmp29(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp29(site) = tmp27(site) - tmp28(site)
        End If
    Next site

    ' 54.åvéZéÆï]âø.HLB_LC_ZLT
    Dim tmp30(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp30(site) = (tmp27(site) + tmp28(site)) / 2
        End If
    Next site

    ' 55.åvéZéÆï]âø.HLB_LC_ZLT
    Dim tmp31(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp31(site) = Div(tmp29(site), tmp30(site), 999)
        End If
    Next site

    ' 56.PutTestResult.HLB_LC_ZLT
    Call ResultAdd("HLB_LC_ZLT", tmp31)

' #### HLB_LC_ZCT ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HLB_LC_ZCT

    ' 1.Clamp.HLB_LC_ZCT

    ' 2.Median.HLB_LC_ZCT

    ' 3.Median.HLB_LC_ZCT

    ' 50.Average_FA.HLB_LC_ZCT
    Dim tmp32_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONECT", EEE_COLOR_ALL, tmp32_0)
    Dim tmp33 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp33, tmp32_0)

    ' 51.GetAverage_Color.HLB_LC_ZCT
    Dim tmp34(nSite) As Double
    Call GetAverage_Color(tmp34, tmp33, "Gr1", "Gr2")

    ' 52.GetAverage_Color.HLB_LC_ZCT
    Dim tmp35(nSite) As Double
    Call GetAverage_Color(tmp35, tmp33, "Gb1", "Gb2")

    ' 53.åvéZéÆï]âø.HLB_LC_ZCT
    Dim tmp36(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp36(site) = tmp34(site) - tmp35(site)
        End If
    Next site

    ' 54.åvéZéÆï]âø.HLB_LC_ZCT
    Dim tmp37(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp37(site) = (tmp34(site) + tmp35(site)) / 2
        End If
    Next site

    ' 55.åvéZéÆï]âø.HLB_LC_ZCT
    Dim tmp38(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp38(site) = Div(tmp36(site), tmp37(site), 999)
        End If
    Next site

    ' 56.PutTestResult.HLB_LC_ZCT
    Call ResultAdd("HLB_LC_ZCT", tmp38)

' #### HLB_LC_ZRT ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HLB_LC_ZRT

    ' 1.Clamp.HLB_LC_ZRT

    ' 2.Median.HLB_LC_ZRT

    ' 3.Median.HLB_LC_ZRT

    ' 50.Average_FA.HLB_LC_ZRT
    Dim tmp39_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONERT", EEE_COLOR_ALL, tmp39_0)
    Dim tmp40 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp40, tmp39_0)

    ' 51.GetAverage_Color.HLB_LC_ZRT
    Dim tmp41(nSite) As Double
    Call GetAverage_Color(tmp41, tmp40, "Gr1", "Gr2")

    ' 52.GetAverage_Color.HLB_LC_ZRT
    Dim tmp42(nSite) As Double
    Call GetAverage_Color(tmp42, tmp40, "Gb1", "Gb2")

    ' 53.åvéZéÆï]âø.HLB_LC_ZRT
    Dim tmp43(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp43(site) = tmp41(site) - tmp42(site)
        End If
    Next site

    ' 54.åvéZéÆï]âø.HLB_LC_ZRT
    Dim tmp44(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp44(site) = (tmp41(site) + tmp42(site)) / 2
        End If
    Next site

    ' 55.åvéZéÆï]âø.HLB_LC_ZRT
    Dim tmp45(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp45(site) = Div(tmp43(site), tmp44(site), 999)
        End If
    Next site

    ' 56.PutTestResult.HLB_LC_ZRT
    Call ResultAdd("HLB_LC_ZRT", tmp45)

' #### HLB_LC_ZLC ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HLB_LC_ZLC

    ' 1.Clamp.HLB_LC_ZLC

    ' 2.Median.HLB_LC_ZLC

    ' 3.Median.HLB_LC_ZLC

    ' 50.Average_FA.HLB_LC_ZLC
    Dim tmp46_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONELC", EEE_COLOR_ALL, tmp46_0)
    Dim tmp47 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp47, tmp46_0)

    ' 51.GetAverage_Color.HLB_LC_ZLC
    Dim tmp48(nSite) As Double
    Call GetAverage_Color(tmp48, tmp47, "Gr1", "Gr2")

    ' 52.GetAverage_Color.HLB_LC_ZLC
    Dim tmp49(nSite) As Double
    Call GetAverage_Color(tmp49, tmp47, "Gb1", "Gb2")

    ' 53.åvéZéÆï]âø.HLB_LC_ZLC
    Dim tmp50(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp50(site) = tmp48(site) - tmp49(site)
        End If
    Next site

    ' 54.åvéZéÆï]âø.HLB_LC_ZLC
    Dim tmp51(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp51(site) = (tmp48(site) + tmp49(site)) / 2
        End If
    Next site

    ' 55.åvéZéÆï]âø.HLB_LC_ZLC
    Dim tmp52(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp52(site) = Div(tmp50(site), tmp51(site), 999)
        End If
    Next site

    ' 56.PutTestResult.HLB_LC_ZLC
    Call ResultAdd("HLB_LC_ZLC", tmp52)

' #### HLB_LC_ZCC ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HLB_LC_ZCC

    ' 1.Clamp.HLB_LC_ZCC

    ' 2.Median.HLB_LC_ZCC

    ' 3.Median.HLB_LC_ZCC

    ' 50.Average_FA.HLB_LC_ZCC
    Dim tmp53_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONECC", EEE_COLOR_ALL, tmp53_0)
    Dim tmp54 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp54, tmp53_0)

    ' 51.GetAverage_Color.HLB_LC_ZCC
    Dim tmp55(nSite) As Double
    Call GetAverage_Color(tmp55, tmp54, "Gr1", "Gr2")

    ' 52.GetAverage_Color.HLB_LC_ZCC
    Dim tmp56(nSite) As Double
    Call GetAverage_Color(tmp56, tmp54, "Gb1", "Gb2")

    ' 53.åvéZéÆï]âø.HLB_LC_ZCC
    Dim tmp57(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp57(site) = tmp55(site) - tmp56(site)
        End If
    Next site

    ' 54.åvéZéÆï]âø.HLB_LC_ZCC
    Dim tmp58(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp58(site) = (tmp55(site) + tmp56(site)) / 2
        End If
    Next site

    ' 55.åvéZéÆï]âø.HLB_LC_ZCC
    Dim tmp59(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp59(site) = Div(tmp57(site), tmp58(site), 999)
        End If
    Next site

    ' 56.PutTestResult.HLB_LC_ZCC
    Call ResultAdd("HLB_LC_ZCC", tmp59)

' #### HLB_LC_ZRC ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HLB_LC_ZRC

    ' 1.Clamp.HLB_LC_ZRC

    ' 2.Median.HLB_LC_ZRC

    ' 3.Median.HLB_LC_ZRC

    ' 50.Average_FA.HLB_LC_ZRC
    Dim tmp60_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONERC", EEE_COLOR_ALL, tmp60_0)
    Dim tmp61 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp61, tmp60_0)

    ' 51.GetAverage_Color.HLB_LC_ZRC
    Dim tmp62(nSite) As Double
    Call GetAverage_Color(tmp62, tmp61, "Gr1", "Gr2")

    ' 52.GetAverage_Color.HLB_LC_ZRC
    Dim tmp63(nSite) As Double
    Call GetAverage_Color(tmp63, tmp61, "Gb1", "Gb2")

    ' 53.åvéZéÆï]âø.HLB_LC_ZRC
    Dim tmp64(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp64(site) = tmp62(site) - tmp63(site)
        End If
    Next site

    ' 54.åvéZéÆï]âø.HLB_LC_ZRC
    Dim tmp65(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp65(site) = (tmp62(site) + tmp63(site)) / 2
        End If
    Next site

    ' 55.åvéZéÆï]âø.HLB_LC_ZRC
    Dim tmp66(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp66(site) = Div(tmp64(site), tmp65(site), 999)
        End If
    Next site

    ' 56.PutTestResult.HLB_LC_ZRC
    Call ResultAdd("HLB_LC_ZRC", tmp66)

' #### HLB_LC_ZLB ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HLB_LC_ZLB

    ' 1.Clamp.HLB_LC_ZLB

    ' 2.Median.HLB_LC_ZLB

    ' 3.Median.HLB_LC_ZLB

    ' 50.Average_FA.HLB_LC_ZLB
    Dim tmp67_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONELB", EEE_COLOR_ALL, tmp67_0)
    Dim tmp68 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp68, tmp67_0)

    ' 51.GetAverage_Color.HLB_LC_ZLB
    Dim tmp69(nSite) As Double
    Call GetAverage_Color(tmp69, tmp68, "Gr1", "Gr2")

    ' 52.GetAverage_Color.HLB_LC_ZLB
    Dim tmp70(nSite) As Double
    Call GetAverage_Color(tmp70, tmp68, "Gb1", "Gb2")

    ' 53.åvéZéÆï]âø.HLB_LC_ZLB
    Dim tmp71(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp71(site) = tmp69(site) - tmp70(site)
        End If
    Next site

    ' 54.åvéZéÆï]âø.HLB_LC_ZLB
    Dim tmp72(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp72(site) = (tmp69(site) + tmp70(site)) / 2
        End If
    Next site

    ' 55.åvéZéÆï]âø.HLB_LC_ZLB
    Dim tmp73(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp73(site) = Div(tmp71(site), tmp72(site), 999)
        End If
    Next site

    ' 56.PutTestResult.HLB_LC_ZLB
    Call ResultAdd("HLB_LC_ZLB", tmp73)

' #### HLB_LC_ZCB ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HLB_LC_ZCB

    ' 1.Clamp.HLB_LC_ZCB

    ' 2.Median.HLB_LC_ZCB

    ' 3.Median.HLB_LC_ZCB

    ' 50.Average_FA.HLB_LC_ZCB
    Dim tmp74_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONECB", EEE_COLOR_ALL, tmp74_0)
    Dim tmp75 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp75, tmp74_0)

    ' 51.GetAverage_Color.HLB_LC_ZCB
    Dim tmp76(nSite) As Double
    Call GetAverage_Color(tmp76, tmp75, "Gr1", "Gr2")

    ' 52.GetAverage_Color.HLB_LC_ZCB
    Dim tmp77(nSite) As Double
    Call GetAverage_Color(tmp77, tmp75, "Gb1", "Gb2")

    ' 53.åvéZéÆï]âø.HLB_LC_ZCB
    Dim tmp78(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp78(site) = tmp76(site) - tmp77(site)
        End If
    Next site

    ' 54.åvéZéÆï]âø.HLB_LC_ZCB
    Dim tmp79(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp79(site) = (tmp76(site) + tmp77(site)) / 2
        End If
    Next site

    ' 55.åvéZéÆï]âø.HLB_LC_ZCB
    Dim tmp80(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp80(site) = Div(tmp78(site), tmp79(site), 999)
        End If
    Next site

    ' 56.PutTestResult.HLB_LC_ZCB
    Call ResultAdd("HLB_LC_ZCB", tmp80)

' #### HLB_LC_ZRB ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HLB_LC_ZRB

    ' 1.Clamp.HLB_LC_ZRB

    ' 2.Median.HLB_LC_ZRB

    ' 3.Median.HLB_LC_ZRB

    ' 50.Average_FA.HLB_LC_ZRB
    Dim tmp81_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONERB", EEE_COLOR_ALL, tmp81_0)
        Call ReleasePlane(sPlane3)
    Dim tmp82 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp82, tmp81_0)

    ' 51.GetAverage_Color.HLB_LC_ZRB
    Dim tmp83(nSite) As Double
    Call GetAverage_Color(tmp83, tmp82, "Gr1", "Gr2")

    ' 52.GetAverage_Color.HLB_LC_ZRB
    Dim tmp84(nSite) As Double
    Call GetAverage_Color(tmp84, tmp82, "Gb1", "Gb2")

    ' 53.åvéZéÆï]âø.HLB_LC_ZRB
    Dim tmp85(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp85(site) = tmp83(site) - tmp84(site)
        End If
    Next site

    ' 54.åvéZéÆï]âø.HLB_LC_ZRB
    Dim tmp86(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp86(site) = (tmp83(site) + tmp84(site)) / 2
        End If
    Next site

    ' 55.åvéZéÆï]âø.HLB_LC_ZRB
    Dim tmp87(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp87(site) = Div(tmp85(site), tmp86(site), 999)
        End If
    Next site

    ' 56.PutTestResult.HLB_LC_ZRB
    Call ResultAdd("HLB_LC_ZRB", tmp87)

End Function



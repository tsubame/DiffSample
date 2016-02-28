Attribute VB_Name = "Image_027_HL3H_ERR_Mod"

Option Explicit

Public Function HL3H_ERR_Process()

        Call PutImageInto_Common

' #### HL3H_SRLNRG ####

    Dim site As Long

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL3H_SRLNRG
    Dim HL3H_ERR_Param As CParamPlane
    Dim HL3H_ERR_DevInfo As CDeviceConfigInfo
    Dim HL3H_ERR_Plane As CImgPlane
    Set HL3H_ERR_Param = TheParameterBank.Item("HL3HImageTest_Acq1")
    Set HL3H_ERR_DevInfo = TheDeviceProfiler.ConfigInfo("HL3HImageTest_Acq1")
        Call TheParameterBank.Delete("HL3HImageTest_Acq1")
    Set HL3H_ERR_Plane = HL3H_ERR_Param.plane

    ' 1.Median.HL3H_SRLNRG
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call MedianEx(HL3H_ERR_Plane, sPlane1, "Bayer2x4_ZONE3", 1, 5)

    ' 2.Median.HL3H_SRLNRG
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane1)

    ' 4.Average_FA.HL3H_SRLNRG
    Dim tmp1_0 As CImgColorAllResult
    Call Average_FA(sPlane2, "Bayer2x4_ZONE22", EEE_COLOR_ALL, tmp1_0)
        Call ReleasePlane(sPlane2)
    Dim tmp2 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp2, tmp1_0)

    ' 5.GetAverage_Color.HL3H_SRLNRG
    Dim tmp3(nSite) As Double
    Call GetAverage_Color(tmp3, tmp2, "R1", "R2")

    ' 6.GetAverage_Color.HL3H_SRLNRG
    Dim tmp4(nSite) As Double
    Call GetAverage_Color(tmp4, tmp2, "Gr1", "Gr2")

    ' 8.ÉfÅ[É^Clamp.HL3H_SRLNRG
    Dim tmp5(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp5(site) = 0
        End If
    Next site
    Dim tmp6 As CImgColorAllResult
    Call Average_FA(HL3H_ERR_Plane, "Bayer2x4_VOPB", EEE_COLOR_ALL, tmp6)
    Call GetAverage_Color(tmp5, tmp6, "-")

    ' 9.åvéZéÆï]âø.HL3H_SRLNRG
    Dim tmp7(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp7(site) = tmp4(site) - tmp5(site)
        End If
    Next site

    ' 11.åvéZéÆï]âø.HL3H_SRLNRG
    Dim tmp8(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp8(site) = tmp3(site) - tmp5(site)
        End If
    Next site

    ' 12.åvéZéÆï]âø.HL3H_SRLNRG
    Dim tmp9(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp9(site) = Div(tmp8(site), tmp7(site), 999)
        End If
    Next site

    ' 14.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL3H_SRLNRG
    Dim HL_MGERR_Param As CParamPlane
    Dim HL_MGERR_DevInfo As CDeviceConfigInfo
    Dim HL_MGERR_Plane As CImgPlane
    Set HL_MGERR_Param = TheParameterBank.Item("HL_MGImageTest_Acq1")
    Set HL_MGERR_DevInfo = TheDeviceProfiler.ConfigInfo("HL_MGImageTest_Acq1")
        Call TheParameterBank.Delete("HL_MGImageTest_Acq1")
    Set HL_MGERR_Plane = HL_MGERR_Param.plane

    ' 15.Median.HL3H_SRLNRG
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call MedianEx(HL_MGERR_Plane, sPlane3, "Bayer2x4_ZONE3", 1, 5)

    ' 16.Median.HL3H_SRLNRG
    Dim sPlane4 As CImgPlane
    Call GetFreePlane(sPlane4, "Normal_Bayer2x4", idpDepthS16, , "sPlane4")
    Call MedianEx(sPlane3, sPlane4, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane3)

    ' 18.Average_FA.HL3H_SRLNRG
    Dim tmp10_0 As CImgColorAllResult
    Call Average_FA(sPlane4, "Bayer2x4_ZONE22", EEE_COLOR_ALL, tmp10_0)
        Call ReleasePlane(sPlane4)
    Dim tmp11 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp11, tmp10_0)

    ' 19.GetAverage_Color.HL3H_SRLNRG
    Dim tmp12(nSite) As Double
    Call GetAverage_Color(tmp12, tmp11, "R1", "R2")

    ' 21.GetAverage_Color.HL3H_SRLNRG
    Dim tmp13(nSite) As Double
    Call GetAverage_Color(tmp13, tmp11, "Gr1", "Gr2")

    ' 22.ÉfÅ[É^Clamp.HL3H_SRLNRG
    Dim tmp14(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp14(site) = 0
        End If
    Next site
    Dim tmp15 As CImgColorAllResult
    Call Average_FA(HL_MGERR_Plane, "Bayer2x4_VOPB", EEE_COLOR_ALL, tmp15)
    Call GetAverage_Color(tmp14, tmp15, "-")

    ' 23.åvéZéÆï]âø.HL3H_SRLNRG
    Dim tmp16(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp16(site) = tmp13(site) - tmp14(site)
        End If
    Next site

    ' 24.åvéZéÆï]âø.HL3H_SRLNRG
    Dim tmp17(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp17(site) = tmp12(site) - tmp14(site)
        End If
    Next site

    ' 28.åvéZéÆï]âø.HL3H_SRLNRG
    Dim tmp18(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp18(site) = Div(tmp17(site), tmp16(site), 999)
        End If
    Next site

    ' 29.åvéZéÆï]âø.HL3H_SRLNRG
    Dim tmp19(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp19(site) = IIf(tmp9(site) = 999 Or tmp18(site) = 999, 999, Div(tmp9(site), tmp18(site), 999))
        End If
    Next site

    ' 30.PutTestResult.HL3H_SRLNRG
    Call ResultAdd("HL3H_SRLNRG", tmp19)

' #### HL3H_SRLNBG ####

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL3H_SRLNBG

    ' 1.Median.HL3H_SRLNBG

    ' 2.Median.HL3H_SRLNBG

    ' 4.Average_FA.HL3H_SRLNBG

    ' 5.GetAverage_Color.HL3H_SRLNBG
    Dim tmp20(nSite) As Double
    Call GetAverage_Color(tmp20, tmp2, "B1", "B2")

    ' 6.GetAverage_Color.HL3H_SRLNBG
    Dim tmp21(nSite) As Double
    Call GetAverage_Color(tmp21, tmp2, "Gb1", "Gb2")

    ' 8.ÉfÅ[É^Clamp.HL3H_SRLNBG

    ' 9.åvéZéÆï]âø.HL3H_SRLNBG
    Dim tmp22(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp22(site) = tmp21(site) - tmp5(site)
        End If
    Next site

    ' 11.åvéZéÆï]âø.HL3H_SRLNBG
    Dim tmp23(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp23(site) = tmp20(site) - tmp5(site)
        End If
    Next site

    ' 12.åvéZéÆï]âø.HL3H_SRLNBG
    Dim tmp24(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp24(site) = Div(tmp23(site), tmp22(site), 999)
        End If
    Next site

    ' 14.âÊëúèÓïÒÉCÉìÉ|Å[Ég.HL3H_SRLNBG

    ' 15.Median.HL3H_SRLNBG

    ' 16.Median.HL3H_SRLNBG

    ' 18.Average_FA.HL3H_SRLNBG

    ' 19.GetAverage_Color.HL3H_SRLNBG
    Dim tmp25(nSite) As Double
    Call GetAverage_Color(tmp25, tmp11, "B1", "B2")

    ' 21.GetAverage_Color.HL3H_SRLNBG
    Dim tmp26(nSite) As Double
    Call GetAverage_Color(tmp26, tmp11, "Gb1", "Gb2")

    ' 22.ÉfÅ[É^Clamp.HL3H_SRLNBG
    Dim tmp27(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp27(site) = 0
        End If
    Next site
    Call GetAverage_Color(tmp27, tmp15, "-")

    ' 23.åvéZéÆï]âø.HL3H_SRLNBG
    Dim tmp28(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp28(site) = tmp26(site) - tmp27(site)
        End If
    Next site

    ' 24.åvéZéÆï]âø.HL3H_SRLNBG
    Dim tmp29(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp29(site) = tmp25(site) - tmp27(site)
        End If
    Next site

    ' 28.åvéZéÆï]âø.HL3H_SRLNBG
    Dim tmp30(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp30(site) = Div(tmp29(site), tmp28(site), 999)
        End If
    Next site

    ' 29.åvéZéÆï]âø.HL3H_SRLNBG
    Dim tmp31(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp31(site) = IIf(tmp24(site) = 999 Or tmp30(site) = 999, 999, Div(tmp24(site), tmp30(site), 999))
        End If
    Next site

    ' 30.PutTestResult.HL3H_SRLNBG
    Call ResultAdd("HL3H_SRLNBG", tmp31)

End Function



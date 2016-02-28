Attribute VB_Name = "Image_019_DKM6_RNERR_Mod"

Option Explicit

Public Function DKM6_RNERR_Process1()

        Call PutImageInto_Common

End Function

Public Function DKM6_RNERR_Process2()

        Call PutImageInto_Common

' #### DK_RNL1_S2 ####

    Dim site As Long

    ' 0.ï°êîâÊëúèÓïÒÉCÉìÉ|Å[Ég.DK_RNL1_S2
    Dim DKM6_RNERR_0_Param As CParamPlane
    Dim DKM6_RNERR_0_DevInfo As CDeviceConfigInfo
    Dim DKM6_RNERR_0_Plane As CImgPlane
    Set DKM6_RNERR_0_Param = TheParameterBank.Item("DKM6_RNImageTest1_Acq1")
    Set DKM6_RNERR_0_DevInfo = TheDeviceProfiler.ConfigInfo("DKM6_RNImageTest1_Acq1")
        Call TheParameterBank.Delete("DKM6_RNImageTest1_Acq1")
    Set DKM6_RNERR_0_Plane = DKM6_RNERR_0_Param.plane

    ' 1.ï°êîâÊëúèÓïÒÉCÉìÉ|Å[Ég.DK_RNL1_S2
    Dim DKM6_RNERR_1_Param As CParamPlane
    Dim DKM6_RNERR_1_DevInfo As CDeviceConfigInfo
    Dim DKM6_RNERR_1_Plane As CImgPlane
    Set DKM6_RNERR_1_Param = TheParameterBank.Item("DKM6_RNImageTest2_Acq1")
    Set DKM6_RNERR_1_DevInfo = TheDeviceProfiler.ConfigInfo("DKM6_RNImageTest2_Acq1")
        Call TheParameterBank.Delete("DKM6_RNImageTest2_Acq1")
    Set DKM6_RNERR_1_Plane = DKM6_RNERR_1_Param.plane

    ' 2.Subtract(í èÌ).DK_RNL1_S2
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call Subtract(DKM6_RNERR_0_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, DKM6_RNERR_1_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane1, "Bayer2x4_FULL", EEE_COLOR_ALL)

    ' 3.ExecuteLUT.DK_RNL1_S2
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call ExecuteLUT(sPlane1, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane2, "Bayer2x4_FULL", EEE_COLOR_ALL, "lut_2")
        Call ReleasePlane(sPlane1)

    ' 8.ShiftLeft.DK_RNL1_S2
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call ShiftLeft(sPlane2, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, 1)
        Call ReleasePlane(sPlane2)

    ' 9.ï°êîâÊëúóp_LSBíËã`.DK_RNL1_S2
    Dim DKM6_RNERR_LSB() As Double
     DKM6_RNERR_LSB = DKM6_RNERR_0_DevInfo.Lsb.AsDouble

    ' 12.SliceLevelê∂ê¨.DK_RNL1_S2
    Dim tmp_Slice1(nSite) As Double
    Call MakeSliceLevel(tmp_Slice1, 0.095 * Sqr(2) * 2 ^ 1, DKM6_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNL1_S2
    Dim tmp1_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice1, tmp_Slice1, idpLimitEachSite, idpLimitExclude, tmp1_0)
        Call ReleasePlane(sPlane3)
    Dim tmp2 As CImgColorAllResult
    Call GetSum_CImgColor(tmp2, tmp1_0)

    ' 14.GetSum_Color.DK_RNL1_S2
    Dim tmp3(nSite) As Double
    Call GetSum_Color(tmp3, tmp2, "-")

    ' 15.PutTestResult.DK_RNL1_S2
    Call ResultAdd("DK_RNL1_S2", tmp3)

' #### DK_RNL1 ####

    ' 0.ë™íËåãâ éÊìæ.DK_RNL1
    Dim tmp_DK_RNL1_S1() As Double
    TheResult.GetResult "DK_RNL1_S1", tmp_DK_RNL1_S1

    ' 1.ë™íËåãâ éÊìæ.DK_RNL1
    Dim tmp_DK_RNL1_S2() As Double
    TheResult.GetResult "DK_RNL1_S2", tmp_DK_RNL1_S2

    ' 2.åvéZéÆï]âø.DK_RNL1
    Dim tmp4(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp4(site) = tmp_DK_RNL1_S1(site) - tmp_DK_RNL1_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNL1
    Call ResultAdd("DK_RNL1", tmp4)

' #### DK_RNL1_1M ####

    ' 0.GetResult.DK_RNL1_1M
    Dim tmp_DK_RNL1() As Double
     TheResult.GetResult "DK_RNL1", tmp_DK_RNL1

    ' 1.åvéZéÆï]âø.DK_RNL1_1M
    Dim tmp5(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp5(site) = TheIDP.PMD("Bayer2x4_ZONE2D").width
        End If
    Next site

    ' 2.åvéZéÆï]âø.DK_RNL1_1M
    Dim tmp6(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp6(site) = TheIDP.PMD("Bayer2x4_ZONE2D").height
        End If
    Next site

    ' 3.åvéZéÆï]âø.DK_RNL1_1M
    Dim tmp7(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp7(site) = 1000000
        End If
    Next site

    ' 4.åvéZéÆï]âø.DK_RNL1_1M
    Dim tmp8(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp8(site) = Div(tmp7(site), tmp5(site) * tmp6(site), 999)
        End If
    Next site

    ' 5.åvéZéÆï]âø.DK_RNL1_1M
    Dim tmp9(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp9(site) = tmp8(site) * tmp_DK_RNL1(site)
        End If
    Next site

    ' 6.PutTestResult.DK_RNL1_1M
    Call ResultAdd("DK_RNL1_1M", tmp9)

End Function



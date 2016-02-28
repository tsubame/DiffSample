Attribute VB_Name = "Image_025_DKH_MGERR_Mod"

Option Explicit

Public Function DKH_MGERR_Process()

        Call PutImageInto_Common

' #### DKH_FDL ####

    Dim site As Long

    ' 0.âÊëúèÓïÒÉCÉìÉ|Å[Ég.DKH_FDL
    Dim DKH_MGERR_Param As CParamPlane
    Dim DKH_MGERR_DevInfo As CDeviceConfigInfo
    Dim DKH_MGERR_Plane As CImgPlane
    Set DKH_MGERR_Param = TheParameterBank.Item("DKH_MGImageTest_Acq1")
    Set DKH_MGERR_DevInfo = TheDeviceProfiler.ConfigInfo("DKH_MGImageTest_Acq1")
        Call TheParameterBank.Delete("DKH_MGImageTest_Acq1")
    Set DKH_MGERR_Plane = DKH_MGERR_Param.plane

    ' 2.Median.DKH_FDL
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call MedianEx(DKH_MGERR_Plane, sPlane1, "Bayer2x4_ZONE3", 1, 7)
    Call MedianEx(DKH_MGERR_Plane, sPlane1, "Bayer2x4_VOPB", 1, 7)

    ' 3.Median.DKH_FDL
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 7, 1)
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_VOPB", 7, 1)
        Call ReleasePlane(sPlane1)

    ' 4.Subtract(í èÌ).DKH_FDL
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call Subtract(DKH_MGERR_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane2, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane3, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane2)

    ' 5.LSBíËã`.DKH_FDL
    Dim DKH_MGERR_LSB() As Double
     DKH_MGERR_LSB = DKH_MGERR_DevInfo.Lsb.AsDouble

    ' 6.SliceLevelê∂ê¨.DKH_FDL
    Dim tmp_Slice1(nSite) As Double
    Call MakeSliceLevel(tmp_Slice1, 0.004, DKH_MGERR_LSB, , , , idpCountAbove)

    ' 8.FDã§óLCopyMask.DKH_FDL
    Dim sPlane4 As CImgPlane
    Call GetFreePlane(sPlane4, "Normal_Bayer2x4", idpDepthS16, True, "sPlane4")
    Call Copy(sPlane3, sPlane3.BasePMD.Name, EEE_COLOR_FLAT, sPlane4, sPlane4.BasePMD.Name, EEE_COLOR_FLAT)
        Call ReleasePlane(sPlane3)

    ' 9.Multimean.DKH_FDL
    Dim sPlane5 As CImgPlane
    Call GetFreePlane(sPlane5, "Normal_Bayer2x4_MUL", idpDepthS16, , "sPlane5")
    Call MakeMulPMD(sPlane5, "Bayer2x4_ZONE2D", "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", 2, 4, EEE_COLOR_FLAT)
    Call MultiMean(sPlane4, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane5, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", EEE_COLOR_FLAT, idpMultiMeanFuncMin, 2, 4)
    Call MakeMulPMD(sPlane5, "Bayer2x4_VOPB", "Bayer2x4_MUL_VOPB_EEE_COLOR_FLAT_2_4", 2, 4, EEE_COLOR_FLAT)
    Call MultiMean(sPlane4, "Bayer2x4_VOPB", EEE_COLOR_FLAT, sPlane5, "Bayer2x4_MUL_VOPB_EEE_COLOR_FLAT_2_4", EEE_COLOR_FLAT, idpMultiMeanFuncMin, 2, 4)

    ' 10.Multimean.DKH_FDL
    Dim sPlane6 As CImgPlane
    Call GetFreePlane(sPlane6, "Normal_Bayer2x4_MUL", idpDepthS16, , "sPlane6")
    Call MakeMulPMD(sPlane6, "Bayer2x4_ZONE2D", "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", 2, 4, EEE_COLOR_FLAT)
    Call MultiMean(sPlane4, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane6, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", EEE_COLOR_FLAT, idpMultiMeanFuncMax, 2, 4)
    Call MakeMulPMD(sPlane6, "Bayer2x4_VOPB", "Bayer2x4_MUL_VOPB_EEE_COLOR_FLAT_2_4", 2, 4, EEE_COLOR_FLAT)
    Call MultiMean(sPlane4, "Bayer2x4_VOPB", EEE_COLOR_FLAT, sPlane6, "Bayer2x4_MUL_VOPB_EEE_COLOR_FLAT_2_4", EEE_COLOR_FLAT, idpMultiMeanFuncMax, 2, 4)
        Call ReleasePlane(sPlane4)

    ' 11.åvéZéÆï]âø.DKH_FDL
    Dim tmp1(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp1(site) = 1
        End If
    Next site

    ' 12.åvéZéÆï]âø.DKH_FDL
    Dim tmp2(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp2(site) = tmp1(site) / 1000
        End If
    Next site

    ' 13.SliceLevelê∂ê¨.DKH_FDL
    Dim tmp_Slice2(nSite) As Double
    Call MakeSliceLevel(tmp_Slice2, tmp2, DKH_MGERR_LSB, , , , idpCountAbove)

    ' 14.PutFlag_FA.DKH_FDL
    Call PutFlag_FA(sPlane5, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", EEE_COLOR_ALL, idpCountAbove, tmp_Slice2, tmp_Slice2, idpLimitEachSite, idpLimitExclude, "Flg_Temp1")
    Call PutFlag_FA(sPlane5, "Bayer2x4_MUL_VOPB_EEE_COLOR_FLAT_2_4", EEE_COLOR_ALL, idpCountAbove, tmp_Slice2, tmp_Slice2, idpLimitEachSite, idpLimitExclude, "Flg_Temp1")
        Call ReleasePlane(sPlane5)

    ' 16.Count_FA.DKH_FDL
    Dim tmp3_0 As CImgColorAllResult
    Call count_FA(sPlane6, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", EEE_COLOR_ALL, idpCountAbove, tmp_Slice1, tmp_Slice1, idpLimitEachSite, idpLimitExclude, tmp3_0, "Flg_Temp2", "Flg_Temp1")
    Dim tmp3_1 As CImgColorAllResult
    Call count_FA(sPlane6, "Bayer2x4_MUL_VOPB_EEE_COLOR_FLAT_2_4", EEE_COLOR_ALL, idpCountAbove, tmp_Slice1, tmp_Slice1, idpLimitEachSite, idpLimitExclude, tmp3_1, "Flg_Temp2", "Flg_Temp1")
        Call ClearALLFlagBit("Flg_Temp1")
    Dim tmp4 As CImgColorAllResult
    Call GetSum_CImgColor(tmp4, tmp3_0, tmp3_1)

    ' 17.GetSum_Color.DKH_FDL
    Dim tmp5(nSite) As Double
    Call GetSum_Color(tmp5, tmp4, "-")

    ' 18.readPixelSite.DKH_FDL
    Dim tmp_RPD1_0(nSite) As CPixInfo
    Call ReadPixelSite(sPlane6, "Bayer2x4_MUL_ZONE2D_EEE_COLOR_FLAT_2_4", tmp5, "Flg_Temp2", tmp_RPD1_0, idpAddrAbsolute)
    Dim tmp_RPD1_1(nSite) As CPixInfo
    Call ReadPixelSite(sPlane6, "Bayer2x4_MUL_VOPB_EEE_COLOR_FLAT_2_4", tmp5, "Flg_Temp2", tmp_RPD1_1, idpAddrAbsolute)
        Call ClearALLFlagBit("Flg_Temp2")
        Call ReleasePlane(sPlane6)
    Dim tmp_RPD2(nSite) As CPixInfo
    Call RPDUnion(tmp_RPD2, tmp_RPD1_0, tmp_RPD1_1)

    ' 19.RPDOffset.DKH_FDL
    Dim tmp_RPD3(nSite) As CPixInfo
    Call RPDOffset(tmp_RPD3, tmp_RPD2, -(2 - 1) + (1 - 1), -(4 - 1) + (3 - 1), 2, 4, 1)

    ' 25.PutTestResult.DKH_FDL
    Call ResultAdd("DKH_FDL", tmp5)

End Function



Attribute VB_Name = "Image_017_DK_RVLERR_VS_Mod"

Option Explicit

Public Function DK_RVLERR_VS_Process1()

        Call PutImageInto_Common

End Function

Public Function DK_RVLERR_VS_Process2()

        Call PutImageInto_Common

' #### DK_RVLV14_VS_S1 ####

    Dim site As Long

    ' 0.複数画像情報インポート.DK_RVLV14_VS_S1
    Dim DK_RVLERR_VS_0_Param As CParamPlane
    Dim DK_RVLERR_VS_0_DevInfo As CDeviceConfigInfo
    Dim DK_RVLERR_VS_0_Plane As CImgPlane
    Set DK_RVLERR_VS_0_Param = TheParameterBank.Item("DK_RVL_VSImageTest1_Acq1")
    Set DK_RVLERR_VS_0_DevInfo = TheDeviceProfiler.ConfigInfo("DK_RVL_VSImageTest1_Acq1")
        Call TheParameterBank.Delete("DK_RVL_VSImageTest1_Acq1")
    Set DK_RVLERR_VS_0_Plane = DK_RVLERR_VS_0_Param.plane

    ' 3.複数画像情報インポート.DK_RVLV14_VS_S1
    Dim DK_RVLERR_VS_1_Param As CParamPlane
    Dim DK_RVLERR_VS_1_DevInfo As CDeviceConfigInfo
    Dim DK_RVLERR_VS_1_Plane As CImgPlane
    Set DK_RVLERR_VS_1_Param = TheParameterBank.Item("DK_RVL_VSImageTest2_Acq1")
    Set DK_RVLERR_VS_1_DevInfo = TheDeviceProfiler.ConfigInfo("DK_RVL_VSImageTest2_Acq1")
        Call TheParameterBank.Delete("DK_RVL_VSImageTest2_Acq1")
    Set DK_RVLERR_VS_1_Plane = DK_RVLERR_VS_1_Param.plane

    ' 6.Subtract(通常).DK_RVLV14_VS_S1
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call Subtract(DK_RVLERR_VS_0_Plane, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, DK_RVLERR_VS_1_Plane, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, sPlane1, "Bayer2x4_ZONE2D", EEE_COLOR_ALL)

    ' 7.ZONE取得.DK_RVLV14_VS_S1

    ' 9.AccumulateColumn.DK_RVLV14_VS_S1
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4_ACC", idpDepthF32, , "sPlane2")
    Call MakeAccPMD(sPlane2, "Bayer2x4_ZONE2D", "Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT)
    Call AccumulateColumn(sPlane1, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane2, "Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, idpStdDeviation)
        Call ReleasePlane(sPlane1)

    ' 10.複数画像用_LSB定義.DK_RVLV14_VS_S1
    Dim DK_RVLERR_VS_LSB() As Double
     DK_RVLERR_VS_LSB = DK_RVLERR_VS_0_DevInfo.Lsb.AsDouble

    ' 11.計算式評価.DK_RVLV14_VS_S1
    Dim tmp1(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp1(site) = 0.00014 / DK_RVLERR_VS_LSB(site)
        End If
    Next site

    ' 12.Count_FA.DK_RVLV14_VS_S1
    Dim tmp2_0 As CImgColorAllResult
    Call count_FA(sPlane2, "Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_ALL, idpCountAbove, tmp1, tmp1, idpLimitEachSite, idpLimitExclude, tmp2_0, "FLG_DK_RVLV14_VS_S1")
        Call ClearALLFlagBit("FLG_DK_RVLV14_VS_S1")
    Dim tmp3 As CImgColorAllResult
    Call GetSum_CImgColor(tmp3, tmp2_0)

    ' 13.GetSum_Color.DK_RVLV14_VS_S1
    Dim tmp4(nSite) As Double
    Call GetSum_Color(tmp4, tmp3, "-")

    ' 14.PutTestResult.DK_RVLV14_VS_S1
    Call ResultAdd("DK_RVLV14_VS_S1", tmp4)

' #### DK_RVLV14_VS_S2 ####

    ' 0.複数画像情報インポート.DK_RVLV14_VS_S2

    ' 3.複数画像情報インポート.DK_RVLV14_VS_S2

    ' 6.Subtract(通常).DK_RVLV14_VS_S2

    ' 7.ZONE取得.DK_RVLV14_VS_S2

    ' 9.AccumulateColumn.DK_RVLV14_VS_S2

    ' 10.複数画像用_LSB定義.DK_RVLV14_VS_S2

    ' 11.計算式評価.DK_RVLV14_VS_S2
    Dim tmp5(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp5(site) = 0.0003 / DK_RVLERR_VS_LSB(site)
        End If
    Next site

    ' 12.Count_FA.DK_RVLV14_VS_S2
    Dim tmp6_0 As CImgColorAllResult
    Call count_FA(sPlane2, "Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_ALL, idpCountAbove, tmp5, tmp5, idpLimitEachSite, idpLimitExclude, tmp6_0)
    Dim tmp7 As CImgColorAllResult
    Call GetSum_CImgColor(tmp7, tmp6_0)

    ' 13.GetSum_Color.DK_RVLV14_VS_S2
    Dim tmp8(nSite) As Double
    Call GetSum_Color(tmp8, tmp7, "-")

    ' 14.PutTestResult.DK_RVLV14_VS_S2
    Call ResultAdd("DK_RVLV14_VS_S2", tmp8)

' #### DK_RVLV14_VS ####

    ' 0.測定結果取得.DK_RVLV14_VS
    Dim tmp_DK_RVLV14_VS_S1() As Double
    TheResult.GetResult "DK_RVLV14_VS_S1", tmp_DK_RVLV14_VS_S1

    ' 1.測定結果取得.DK_RVLV14_VS
    Dim tmp_DK_RVLV14_VS_S2() As Double
    TheResult.GetResult "DK_RVLV14_VS_S2", tmp_DK_RVLV14_VS_S2

    ' 2.計算式評価.DK_RVLV14_VS
    Dim tmp9(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp9(site) = tmp_DK_RVLV14_VS_S1(site) - tmp_DK_RVLV14_VS_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RVLV14_VS
    Call ResultAdd("DK_RVLV14_VS", tmp9)

' #### DK_RVLV30_VS_S2 ####

    ' 0.複数画像情報インポート.DK_RVLV30_VS_S2

    ' 3.複数画像情報インポート.DK_RVLV30_VS_S2

    ' 6.Subtract(通常).DK_RVLV30_VS_S2

    ' 7.ZONE取得.DK_RVLV30_VS_S2

    ' 9.AccumulateColumn.DK_RVLV30_VS_S2

    ' 10.複数画像用_LSB定義.DK_RVLV30_VS_S2

    ' 11.計算式評価.DK_RVLV30_VS_S2
    Dim tmp10(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp10(site) = 0.00045 / DK_RVLERR_VS_LSB(site)
        End If
    Next site

    ' 12.Count_FA.DK_RVLV30_VS_S2
    Dim tmp11_0 As CImgColorAllResult
    Call count_FA(sPlane2, "Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_ALL, idpCountAbove, tmp10, tmp10, idpLimitEachSite, idpLimitExclude, tmp11_0)
    Dim tmp12 As CImgColorAllResult
    Call GetSum_CImgColor(tmp12, tmp11_0)

    ' 13.GetSum_Color.DK_RVLV30_VS_S2
    Dim tmp13(nSite) As Double
    Call GetSum_Color(tmp13, tmp12, "-")

    ' 14.PutTestResult.DK_RVLV30_VS_S2
    Call ResultAdd("DK_RVLV30_VS_S2", tmp13)

' #### DK_RVLV30_VS ####

    ' 0.測定結果取得.DK_RVLV30_VS

    ' 1.測定結果取得.DK_RVLV30_VS
    Dim tmp_DK_RVLV30_VS_S2() As Double
    TheResult.GetResult "DK_RVLV30_VS_S2", tmp_DK_RVLV30_VS_S2

    ' 2.計算式評価.DK_RVLV30_VS
    Dim tmp14(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp14(site) = tmp_DK_RVLV14_VS_S2(site) - tmp_DK_RVLV30_VS_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RVLV30_VS
    Call ResultAdd("DK_RVLV30_VS", tmp14)

' #### DK_RVLV45_VS_S2 ####

    ' 0.複数画像情報インポート.DK_RVLV45_VS_S2

    ' 3.複数画像情報インポート.DK_RVLV45_VS_S2

    ' 6.Subtract(通常).DK_RVLV45_VS_S2

    ' 7.ZONE取得.DK_RVLV45_VS_S2

    ' 9.AccumulateColumn.DK_RVLV45_VS_S2

    ' 10.複数画像用_LSB定義.DK_RVLV45_VS_S2

    ' 11.計算式評価.DK_RVLV45_VS_S2
    Dim tmp15(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp15(site) = 0.0006 / DK_RVLERR_VS_LSB(site)
        End If
    Next site

    ' 12.Count_FA.DK_RVLV45_VS_S2
    Dim tmp16_0 As CImgColorAllResult
    Call count_FA(sPlane2, "Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_ALL, idpCountAbove, tmp15, tmp15, idpLimitEachSite, idpLimitExclude, tmp16_0)
    Dim tmp17 As CImgColorAllResult
    Call GetSum_CImgColor(tmp17, tmp16_0)

    ' 13.GetSum_Color.DK_RVLV45_VS_S2
    Dim tmp18(nSite) As Double
    Call GetSum_Color(tmp18, tmp17, "-")

    ' 14.PutTestResult.DK_RVLV45_VS_S2
    Call ResultAdd("DK_RVLV45_VS_S2", tmp18)

' #### DK_RVLV45_VS ####

    ' 0.測定結果取得.DK_RVLV45_VS

    ' 1.測定結果取得.DK_RVLV45_VS
    Dim tmp_DK_RVLV45_VS_S2() As Double
    TheResult.GetResult "DK_RVLV45_VS_S2", tmp_DK_RVLV45_VS_S2

    ' 2.計算式評価.DK_RVLV45_VS
    Dim tmp19(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp19(site) = tmp_DK_RVLV30_VS_S2(site) - tmp_DK_RVLV45_VS_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RVLV45_VS
    Call ResultAdd("DK_RVLV45_VS", tmp19)

' #### DK_RVLV75_VS ####

    ' 0.複数画像情報インポート.DK_RVLV75_VS

    ' 3.複数画像情報インポート.DK_RVLV75_VS

    ' 6.Subtract(通常).DK_RVLV75_VS

    ' 7.ZONE取得.DK_RVLV75_VS

    ' 9.AccumulateColumn.DK_RVLV75_VS

    ' 10.複数画像用_LSB定義.DK_RVLV75_VS

    ' 11.計算式評価.DK_RVLV75_VS
    Dim tmp20(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp20(site) = 0.00075 / DK_RVLERR_VS_LSB(site)
        End If
    Next site

    ' 12.Count_FA.DK_RVLV75_VS
    Dim tmp21_0 As CImgColorAllResult
    Call count_FA(sPlane2, "Bayer2x4_ACC_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_ALL, idpCountAbove, tmp20, tmp20, idpLimitEachSite, idpLimitExclude, tmp21_0)
        Call ReleasePlane(sPlane2)
    Dim tmp22 As CImgColorAllResult
    Call GetSum_CImgColor(tmp22, tmp21_0)

    ' 13.GetSum_Color.DK_RVLV75_VS
    Dim tmp23(nSite) As Double
    Call GetSum_Color(tmp23, tmp22, "-")

    ' 14.PutTestResult.DK_RVLV75_VS
    Call ResultAdd("DK_RVLV75_VS", tmp23)

' #### DK_RVLV60_VS ####

    ' 0.測定結果取得.DK_RVLV60_VS

    ' 1.測定結果取得.DK_RVLV60_VS
    Dim tmp_DK_RVLV75_VS() As Double
    TheResult.GetResult "DK_RVLV75_VS", tmp_DK_RVLV75_VS

    ' 2.計算式評価.DK_RVLV60_VS
    Dim tmp24(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp24(site) = tmp_DK_RVLV45_VS_S2(site) - tmp_DK_RVLV75_VS(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RVLV60_VS
    Call ResultAdd("DK_RVLV60_VS", tmp24)

End Function



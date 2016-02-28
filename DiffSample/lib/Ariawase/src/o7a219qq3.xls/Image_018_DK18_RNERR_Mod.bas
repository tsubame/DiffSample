Attribute VB_Name = "Image_018_DK18_RNERR_Mod"

Option Explicit

Public Function DK18_RNERR_Process1()

        Call PutImageInto_Common

End Function

Public Function DK18_RNERR_Process2()

        Call PutImageInto_Common

' #### DK_RNV01_S1 ####

    Dim site As Long

    ' 0.複数画像情報インポート.DK_RNV01_S1
    Dim DK18_RNERR_0_Param As CParamPlane
    Dim DK18_RNERR_0_DevInfo As CDeviceConfigInfo
    Dim DK18_RNERR_0_Plane As CImgPlane
    Set DK18_RNERR_0_Param = TheParameterBank.Item("DK18_RNImageTest1_Acq1")
    Set DK18_RNERR_0_DevInfo = TheDeviceProfiler.ConfigInfo("DK18_RNImageTest1_Acq1")
        Call TheParameterBank.Delete("DK18_RNImageTest1_Acq1")
    Set DK18_RNERR_0_Plane = DK18_RNERR_0_Param.plane

    ' 1.複数画像情報インポート.DK_RNV01_S1
    Dim DK18_RNERR_1_Param As CParamPlane
    Dim DK18_RNERR_1_DevInfo As CDeviceConfigInfo
    Dim DK18_RNERR_1_Plane As CImgPlane
    Set DK18_RNERR_1_Param = TheParameterBank.Item("DK18_RNImageTest2_Acq1")
    Set DK18_RNERR_1_DevInfo = TheDeviceProfiler.ConfigInfo("DK18_RNImageTest2_Acq1")
        Call TheParameterBank.Delete("DK18_RNImageTest2_Acq1")
    Set DK18_RNERR_1_Plane = DK18_RNERR_1_Param.plane

    ' 2.Subtract(通常).DK_RNV01_S1
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call Subtract(DK18_RNERR_0_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, DK18_RNERR_1_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane1, "Bayer2x4_FULL", EEE_COLOR_ALL)

    ' 3.ExecuteLUT.DK_RNV01_S1
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call ExecuteLUT(sPlane1, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane2, "Bayer2x4_FULL", EEE_COLOR_ALL, "lut_2")
        Call ReleasePlane(sPlane1)

    ' 8.ShiftLeft.DK_RNV01_S1
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call ShiftLeft(sPlane2, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, 1)

    ' 9.複数画像用_LSB定義.DK_RNV01_S1
    Dim DK18_RNERR_LSB() As Double
     DK18_RNERR_LSB = DK18_RNERR_0_DevInfo.Lsb.AsDouble

    ' 12.SliceLevel生成.DK_RNV01_S1
    Dim tmp_Slice1(nSite) As Double
    Call MakeSliceLevel(tmp_Slice1, 0.0001 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV01_S1
    Dim tmp1_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice1, tmp_Slice1, idpLimitEachSite, idpLimitExclude, tmp1_0, "FLG_DK_RNV01_S1")
        Call ClearALLFlagBit("FLG_DK_RNV01_S1")
    Dim tmp2 As CImgColorAllResult
    Call GetSum_CImgColor(tmp2, tmp1_0)

    ' 14.GetSum_Color.DK_RNV01_S1
    Dim tmp3(nSite) As Double
    Call GetSum_Color(tmp3, tmp2, "-")

    ' 15.PutTestResult.DK_RNV01_S1
    Call ResultAdd("DK_RNV01_S1", tmp3)

' #### DK_RNV01_S2 ####

    ' 0.複数画像情報インポート.DK_RNV01_S2

    ' 1.複数画像情報インポート.DK_RNV01_S2

    ' 2.Subtract(通常).DK_RNV01_S2

    ' 3.ExecuteLUT.DK_RNV01_S2

    ' 8.ShiftLeft.DK_RNV01_S2

    ' 9.複数画像用_LSB定義.DK_RNV01_S2

    ' 12.SliceLevel生成.DK_RNV01_S2
    Dim tmp_Slice2(nSite) As Double
    Call MakeSliceLevel(tmp_Slice2, 0.0002 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV01_S2
    Dim tmp4_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice2, tmp_Slice2, idpLimitEachSite, idpLimitExclude, tmp4_0)
    Dim tmp5 As CImgColorAllResult
    Call GetSum_CImgColor(tmp5, tmp4_0)

    ' 14.GetSum_Color.DK_RNV01_S2
    Dim tmp6(nSite) As Double
    Call GetSum_Color(tmp6, tmp5, "-")

    ' 15.PutTestResult.DK_RNV01_S2
    Call ResultAdd("DK_RNV01_S2", tmp6)

' #### DK_RNV01 ####

    ' 0.測定結果取得.DK_RNV01
    Dim tmp_DK_RNV01_S1() As Double
    TheResult.GetResult "DK_RNV01_S1", tmp_DK_RNV01_S1

    ' 1.測定結果取得.DK_RNV01
    Dim tmp_DK_RNV01_S2() As Double
    TheResult.GetResult "DK_RNV01_S2", tmp_DK_RNV01_S2

    ' 2.計算式評価.DK_RNV01
    Dim tmp7(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp7(site) = tmp_DK_RNV01_S1(site) - tmp_DK_RNV01_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV01
    Call ResultAdd("DK_RNV01", tmp7)

' #### DK_RNV02_S2 ####

    ' 0.複数画像情報インポート.DK_RNV02_S2

    ' 1.複数画像情報インポート.DK_RNV02_S2

    ' 2.Subtract(通常).DK_RNV02_S2

    ' 3.ExecuteLUT.DK_RNV02_S2

    ' 8.ShiftLeft.DK_RNV02_S2

    ' 9.複数画像用_LSB定義.DK_RNV02_S2

    ' 12.SliceLevel生成.DK_RNV02_S2
    Dim tmp_Slice3(nSite) As Double
    Call MakeSliceLevel(tmp_Slice3, 0.0003 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV02_S2
    Dim tmp8_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice3, tmp_Slice3, idpLimitEachSite, idpLimitExclude, tmp8_0)
    Dim tmp9 As CImgColorAllResult
    Call GetSum_CImgColor(tmp9, tmp8_0)

    ' 14.GetSum_Color.DK_RNV02_S2
    Dim tmp10(nSite) As Double
    Call GetSum_Color(tmp10, tmp9, "-")

    ' 15.PutTestResult.DK_RNV02_S2
    Call ResultAdd("DK_RNV02_S2", tmp10)

' #### DK_RNV02 ####

    ' 0.測定結果取得.DK_RNV02

    ' 1.測定結果取得.DK_RNV02
    Dim tmp_DK_RNV02_S2() As Double
    TheResult.GetResult "DK_RNV02_S2", tmp_DK_RNV02_S2

    ' 2.計算式評価.DK_RNV02
    Dim tmp11(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp11(site) = tmp_DK_RNV01_S2(site) - tmp_DK_RNV02_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV02
    Call ResultAdd("DK_RNV02", tmp11)

' #### DK_RNV03_S2 ####

    ' 0.複数画像情報インポート.DK_RNV03_S2

    ' 1.複数画像情報インポート.DK_RNV03_S2

    ' 2.Subtract(通常).DK_RNV03_S2

    ' 3.ExecuteLUT.DK_RNV03_S2

    ' 8.ShiftLeft.DK_RNV03_S2

    ' 9.複数画像用_LSB定義.DK_RNV03_S2

    ' 12.SliceLevel生成.DK_RNV03_S2
    Dim tmp_Slice4(nSite) As Double
    Call MakeSliceLevel(tmp_Slice4, 0.0004 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV03_S2
    Dim tmp12_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice4, tmp_Slice4, idpLimitEachSite, idpLimitExclude, tmp12_0)
    Dim tmp13 As CImgColorAllResult
    Call GetSum_CImgColor(tmp13, tmp12_0)

    ' 14.GetSum_Color.DK_RNV03_S2
    Dim tmp14(nSite) As Double
    Call GetSum_Color(tmp14, tmp13, "-")

    ' 15.PutTestResult.DK_RNV03_S2
    Call ResultAdd("DK_RNV03_S2", tmp14)

' #### DK_RNV03 ####

    ' 0.測定結果取得.DK_RNV03

    ' 1.測定結果取得.DK_RNV03
    Dim tmp_DK_RNV03_S2() As Double
    TheResult.GetResult "DK_RNV03_S2", tmp_DK_RNV03_S2

    ' 2.計算式評価.DK_RNV03
    Dim tmp15(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp15(site) = tmp_DK_RNV02_S2(site) - tmp_DK_RNV03_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV03
    Call ResultAdd("DK_RNV03", tmp15)

' #### DK_RNV04_S2 ####

    ' 0.複数画像情報インポート.DK_RNV04_S2

    ' 1.複数画像情報インポート.DK_RNV04_S2

    ' 2.Subtract(通常).DK_RNV04_S2

    ' 3.ExecuteLUT.DK_RNV04_S2

    ' 8.ShiftLeft.DK_RNV04_S2

    ' 9.複数画像用_LSB定義.DK_RNV04_S2

    ' 12.SliceLevel生成.DK_RNV04_S2
    Dim tmp_Slice5(nSite) As Double
    Call MakeSliceLevel(tmp_Slice5, 0.0005 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV04_S2
    Dim tmp16_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice5, tmp_Slice5, idpLimitEachSite, idpLimitExclude, tmp16_0)
    Dim tmp17 As CImgColorAllResult
    Call GetSum_CImgColor(tmp17, tmp16_0)

    ' 14.GetSum_Color.DK_RNV04_S2
    Dim tmp18(nSite) As Double
    Call GetSum_Color(tmp18, tmp17, "-")

    ' 15.PutTestResult.DK_RNV04_S2
    Call ResultAdd("DK_RNV04_S2", tmp18)

' #### DK_RNV04 ####

    ' 0.測定結果取得.DK_RNV04

    ' 1.測定結果取得.DK_RNV04
    Dim tmp_DK_RNV04_S2() As Double
    TheResult.GetResult "DK_RNV04_S2", tmp_DK_RNV04_S2

    ' 2.計算式評価.DK_RNV04
    Dim tmp19(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp19(site) = tmp_DK_RNV03_S2(site) - tmp_DK_RNV04_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV04
    Call ResultAdd("DK_RNV04", tmp19)

' #### DK_RNV05_S2 ####

    ' 0.複数画像情報インポート.DK_RNV05_S2

    ' 1.複数画像情報インポート.DK_RNV05_S2

    ' 2.Subtract(通常).DK_RNV05_S2

    ' 3.ExecuteLUT.DK_RNV05_S2

    ' 8.ShiftLeft.DK_RNV05_S2

    ' 9.複数画像用_LSB定義.DK_RNV05_S2

    ' 12.SliceLevel生成.DK_RNV05_S2
    Dim tmp_Slice6(nSite) As Double
    Call MakeSliceLevel(tmp_Slice6, 0.0006 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV05_S2
    Dim tmp20_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice6, tmp_Slice6, idpLimitEachSite, idpLimitExclude, tmp20_0)
    Dim tmp21 As CImgColorAllResult
    Call GetSum_CImgColor(tmp21, tmp20_0)

    ' 14.GetSum_Color.DK_RNV05_S2
    Dim tmp22(nSite) As Double
    Call GetSum_Color(tmp22, tmp21, "-")

    ' 15.PutTestResult.DK_RNV05_S2
    Call ResultAdd("DK_RNV05_S2", tmp22)

' #### DK_RNV05 ####

    ' 0.測定結果取得.DK_RNV05

    ' 1.測定結果取得.DK_RNV05
    Dim tmp_DK_RNV05_S2() As Double
    TheResult.GetResult "DK_RNV05_S2", tmp_DK_RNV05_S2

    ' 2.計算式評価.DK_RNV05
    Dim tmp23(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp23(site) = tmp_DK_RNV04_S2(site) - tmp_DK_RNV05_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV05
    Call ResultAdd("DK_RNV05", tmp23)

' #### DK_RNV06_S2 ####

    ' 0.複数画像情報インポート.DK_RNV06_S2

    ' 1.複数画像情報インポート.DK_RNV06_S2

    ' 2.Subtract(通常).DK_RNV06_S2

    ' 3.ExecuteLUT.DK_RNV06_S2

    ' 8.ShiftLeft.DK_RNV06_S2

    ' 9.複数画像用_LSB定義.DK_RNV06_S2

    ' 12.SliceLevel生成.DK_RNV06_S2
    Dim tmp_Slice7(nSite) As Double
    Call MakeSliceLevel(tmp_Slice7, 0.0007 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV06_S2
    Dim tmp24_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice7, tmp_Slice7, idpLimitEachSite, idpLimitExclude, tmp24_0)
    Dim tmp25 As CImgColorAllResult
    Call GetSum_CImgColor(tmp25, tmp24_0)

    ' 14.GetSum_Color.DK_RNV06_S2
    Dim tmp26(nSite) As Double
    Call GetSum_Color(tmp26, tmp25, "-")

    ' 15.PutTestResult.DK_RNV06_S2
    Call ResultAdd("DK_RNV06_S2", tmp26)

' #### DK_RNV06 ####

    ' 0.測定結果取得.DK_RNV06

    ' 1.測定結果取得.DK_RNV06
    Dim tmp_DK_RNV06_S2() As Double
    TheResult.GetResult "DK_RNV06_S2", tmp_DK_RNV06_S2

    ' 2.計算式評価.DK_RNV06
    Dim tmp27(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp27(site) = tmp_DK_RNV05_S2(site) - tmp_DK_RNV06_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV06
    Call ResultAdd("DK_RNV06", tmp27)

' #### DK_RNV07_S2 ####

    ' 0.複数画像情報インポート.DK_RNV07_S2

    ' 1.複数画像情報インポート.DK_RNV07_S2

    ' 2.Subtract(通常).DK_RNV07_S2

    ' 3.ExecuteLUT.DK_RNV07_S2

    ' 8.ShiftLeft.DK_RNV07_S2

    ' 9.複数画像用_LSB定義.DK_RNV07_S2

    ' 12.SliceLevel生成.DK_RNV07_S2
    Dim tmp_Slice8(nSite) As Double
    Call MakeSliceLevel(tmp_Slice8, 0.0008 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV07_S2
    Dim tmp28_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice8, tmp_Slice8, idpLimitEachSite, idpLimitExclude, tmp28_0)
    Dim tmp29 As CImgColorAllResult
    Call GetSum_CImgColor(tmp29, tmp28_0)

    ' 14.GetSum_Color.DK_RNV07_S2
    Dim tmp30(nSite) As Double
    Call GetSum_Color(tmp30, tmp29, "-")

    ' 15.PutTestResult.DK_RNV07_S2
    Call ResultAdd("DK_RNV07_S2", tmp30)

' #### DK_RNV07 ####

    ' 0.測定結果取得.DK_RNV07

    ' 1.測定結果取得.DK_RNV07
    Dim tmp_DK_RNV07_S2() As Double
    TheResult.GetResult "DK_RNV07_S2", tmp_DK_RNV07_S2

    ' 2.計算式評価.DK_RNV07
    Dim tmp31(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp31(site) = tmp_DK_RNV06_S2(site) - tmp_DK_RNV07_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV07
    Call ResultAdd("DK_RNV07", tmp31)

' #### DK_RNV08_S2 ####

    ' 0.複数画像情報インポート.DK_RNV08_S2

    ' 1.複数画像情報インポート.DK_RNV08_S2

    ' 2.Subtract(通常).DK_RNV08_S2

    ' 3.ExecuteLUT.DK_RNV08_S2

    ' 8.ShiftLeft.DK_RNV08_S2

    ' 9.複数画像用_LSB定義.DK_RNV08_S2

    ' 12.SliceLevel生成.DK_RNV08_S2
    Dim tmp_Slice9(nSite) As Double
    Call MakeSliceLevel(tmp_Slice9, 0.0009 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV08_S2
    Dim tmp32_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice9, tmp_Slice9, idpLimitEachSite, idpLimitExclude, tmp32_0)
    Dim tmp33 As CImgColorAllResult
    Call GetSum_CImgColor(tmp33, tmp32_0)

    ' 14.GetSum_Color.DK_RNV08_S2
    Dim tmp34(nSite) As Double
    Call GetSum_Color(tmp34, tmp33, "-")

    ' 15.PutTestResult.DK_RNV08_S2
    Call ResultAdd("DK_RNV08_S2", tmp34)

' #### DK_RNV08 ####

    ' 0.測定結果取得.DK_RNV08

    ' 1.測定結果取得.DK_RNV08
    Dim tmp_DK_RNV08_S2() As Double
    TheResult.GetResult "DK_RNV08_S2", tmp_DK_RNV08_S2

    ' 2.計算式評価.DK_RNV08
    Dim tmp35(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp35(site) = tmp_DK_RNV07_S2(site) - tmp_DK_RNV08_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV08
    Call ResultAdd("DK_RNV08", tmp35)

' #### DK_RNV09_S2 ####

    ' 0.複数画像情報インポート.DK_RNV09_S2

    ' 1.複数画像情報インポート.DK_RNV09_S2

    ' 2.Subtract(通常).DK_RNV09_S2

    ' 3.ExecuteLUT.DK_RNV09_S2

    ' 8.ShiftLeft.DK_RNV09_S2

    ' 9.複数画像用_LSB定義.DK_RNV09_S2

    ' 12.SliceLevel生成.DK_RNV09_S2
    Dim tmp_Slice10(nSite) As Double
    Call MakeSliceLevel(tmp_Slice10, 0.001 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV09_S2
    Dim tmp36_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice10, tmp_Slice10, idpLimitEachSite, idpLimitExclude, tmp36_0)
    Dim tmp37 As CImgColorAllResult
    Call GetSum_CImgColor(tmp37, tmp36_0)

    ' 14.GetSum_Color.DK_RNV09_S2
    Dim tmp38(nSite) As Double
    Call GetSum_Color(tmp38, tmp37, "-")

    ' 15.PutTestResult.DK_RNV09_S2
    Call ResultAdd("DK_RNV09_S2", tmp38)

' #### DK_RNV09 ####

    ' 0.測定結果取得.DK_RNV09

    ' 1.測定結果取得.DK_RNV09
    Dim tmp_DK_RNV09_S2() As Double
    TheResult.GetResult "DK_RNV09_S2", tmp_DK_RNV09_S2

    ' 2.計算式評価.DK_RNV09
    Dim tmp39(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp39(site) = tmp_DK_RNV08_S2(site) - tmp_DK_RNV09_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV09
    Call ResultAdd("DK_RNV09", tmp39)

' #### DK_RNV10_S2 ####

    ' 0.複数画像情報インポート.DK_RNV10_S2

    ' 1.複数画像情報インポート.DK_RNV10_S2

    ' 2.Subtract(通常).DK_RNV10_S2

    ' 3.ExecuteLUT.DK_RNV10_S2

    ' 8.ShiftLeft.DK_RNV10_S2

    ' 9.複数画像用_LSB定義.DK_RNV10_S2

    ' 12.SliceLevel生成.DK_RNV10_S2
    Dim tmp_Slice11(nSite) As Double
    Call MakeSliceLevel(tmp_Slice11, 0.0011 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV10_S2
    Dim tmp40_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice11, tmp_Slice11, idpLimitEachSite, idpLimitExclude, tmp40_0)
    Dim tmp41 As CImgColorAllResult
    Call GetSum_CImgColor(tmp41, tmp40_0)

    ' 14.GetSum_Color.DK_RNV10_S2
    Dim tmp42(nSite) As Double
    Call GetSum_Color(tmp42, tmp41, "-")

    ' 15.PutTestResult.DK_RNV10_S2
    Call ResultAdd("DK_RNV10_S2", tmp42)

' #### DK_RNV10 ####

    ' 0.測定結果取得.DK_RNV10

    ' 1.測定結果取得.DK_RNV10
    Dim tmp_DK_RNV10_S2() As Double
    TheResult.GetResult "DK_RNV10_S2", tmp_DK_RNV10_S2

    ' 2.計算式評価.DK_RNV10
    Dim tmp43(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp43(site) = tmp_DK_RNV09_S2(site) - tmp_DK_RNV10_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV10
    Call ResultAdd("DK_RNV10", tmp43)

' #### DK_RNV11_S2 ####

    ' 0.複数画像情報インポート.DK_RNV11_S2

    ' 1.複数画像情報インポート.DK_RNV11_S2

    ' 2.Subtract(通常).DK_RNV11_S2

    ' 3.ExecuteLUT.DK_RNV11_S2

    ' 8.ShiftLeft.DK_RNV11_S2

    ' 9.複数画像用_LSB定義.DK_RNV11_S2

    ' 12.SliceLevel生成.DK_RNV11_S2
    Dim tmp_Slice12(nSite) As Double
    Call MakeSliceLevel(tmp_Slice12, 0.0012 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV11_S2
    Dim tmp44_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice12, tmp_Slice12, idpLimitEachSite, idpLimitExclude, tmp44_0)
    Dim tmp45 As CImgColorAllResult
    Call GetSum_CImgColor(tmp45, tmp44_0)

    ' 14.GetSum_Color.DK_RNV11_S2
    Dim tmp46(nSite) As Double
    Call GetSum_Color(tmp46, tmp45, "-")

    ' 15.PutTestResult.DK_RNV11_S2
    Call ResultAdd("DK_RNV11_S2", tmp46)

' #### DK_RNV11 ####

    ' 0.測定結果取得.DK_RNV11

    ' 1.測定結果取得.DK_RNV11
    Dim tmp_DK_RNV11_S2() As Double
    TheResult.GetResult "DK_RNV11_S2", tmp_DK_RNV11_S2

    ' 2.計算式評価.DK_RNV11
    Dim tmp47(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp47(site) = tmp_DK_RNV10_S2(site) - tmp_DK_RNV11_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV11
    Call ResultAdd("DK_RNV11", tmp47)

' #### DK_RNV12_S2 ####

    ' 0.複数画像情報インポート.DK_RNV12_S2

    ' 1.複数画像情報インポート.DK_RNV12_S2

    ' 2.Subtract(通常).DK_RNV12_S2

    ' 3.ExecuteLUT.DK_RNV12_S2

    ' 8.ShiftLeft.DK_RNV12_S2

    ' 9.複数画像用_LSB定義.DK_RNV12_S2

    ' 12.SliceLevel生成.DK_RNV12_S2
    Dim tmp_Slice13(nSite) As Double
    Call MakeSliceLevel(tmp_Slice13, 0.0013 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV12_S2
    Dim tmp48_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice13, tmp_Slice13, idpLimitEachSite, idpLimitExclude, tmp48_0)
    Dim tmp49 As CImgColorAllResult
    Call GetSum_CImgColor(tmp49, tmp48_0)

    ' 14.GetSum_Color.DK_RNV12_S2
    Dim tmp50(nSite) As Double
    Call GetSum_Color(tmp50, tmp49, "-")

    ' 15.PutTestResult.DK_RNV12_S2
    Call ResultAdd("DK_RNV12_S2", tmp50)

' #### DK_RNV12 ####

    ' 0.測定結果取得.DK_RNV12

    ' 1.測定結果取得.DK_RNV12
    Dim tmp_DK_RNV12_S2() As Double
    TheResult.GetResult "DK_RNV12_S2", tmp_DK_RNV12_S2

    ' 2.計算式評価.DK_RNV12
    Dim tmp51(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp51(site) = tmp_DK_RNV11_S2(site) - tmp_DK_RNV12_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV12
    Call ResultAdd("DK_RNV12", tmp51)

' #### DK_RNV13_S2 ####

    ' 0.複数画像情報インポート.DK_RNV13_S2

    ' 1.複数画像情報インポート.DK_RNV13_S2

    ' 2.Subtract(通常).DK_RNV13_S2

    ' 3.ExecuteLUT.DK_RNV13_S2

    ' 8.ShiftLeft.DK_RNV13_S2

    ' 9.複数画像用_LSB定義.DK_RNV13_S2

    ' 12.SliceLevel生成.DK_RNV13_S2
    Dim tmp_Slice14(nSite) As Double
    Call MakeSliceLevel(tmp_Slice14, 0.0014 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV13_S2
    Dim tmp52_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice14, tmp_Slice14, idpLimitEachSite, idpLimitExclude, tmp52_0)
    Dim tmp53 As CImgColorAllResult
    Call GetSum_CImgColor(tmp53, tmp52_0)

    ' 14.GetSum_Color.DK_RNV13_S2
    Dim tmp54(nSite) As Double
    Call GetSum_Color(tmp54, tmp53, "-")

    ' 15.PutTestResult.DK_RNV13_S2
    Call ResultAdd("DK_RNV13_S2", tmp54)

' #### DK_RNV13 ####

    ' 0.測定結果取得.DK_RNV13

    ' 1.測定結果取得.DK_RNV13
    Dim tmp_DK_RNV13_S2() As Double
    TheResult.GetResult "DK_RNV13_S2", tmp_DK_RNV13_S2

    ' 2.計算式評価.DK_RNV13
    Dim tmp55(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp55(site) = tmp_DK_RNV12_S2(site) - tmp_DK_RNV13_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV13
    Call ResultAdd("DK_RNV13", tmp55)

' #### DK_RNV14_S2 ####

    ' 0.複数画像情報インポート.DK_RNV14_S2

    ' 1.複数画像情報インポート.DK_RNV14_S2

    ' 2.Subtract(通常).DK_RNV14_S2

    ' 3.ExecuteLUT.DK_RNV14_S2

    ' 8.ShiftLeft.DK_RNV14_S2

    ' 9.複数画像用_LSB定義.DK_RNV14_S2

    ' 12.SliceLevel生成.DK_RNV14_S2
    Dim tmp_Slice15(nSite) As Double
    Call MakeSliceLevel(tmp_Slice15, 0.0015 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV14_S2
    Dim tmp56_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice15, tmp_Slice15, idpLimitEachSite, idpLimitExclude, tmp56_0)
    Dim tmp57 As CImgColorAllResult
    Call GetSum_CImgColor(tmp57, tmp56_0)

    ' 14.GetSum_Color.DK_RNV14_S2
    Dim tmp58(nSite) As Double
    Call GetSum_Color(tmp58, tmp57, "-")

    ' 15.PutTestResult.DK_RNV14_S2
    Call ResultAdd("DK_RNV14_S2", tmp58)

' #### DK_RNV14 ####

    ' 0.測定結果取得.DK_RNV14

    ' 1.測定結果取得.DK_RNV14
    Dim tmp_DK_RNV14_S2() As Double
    TheResult.GetResult "DK_RNV14_S2", tmp_DK_RNV14_S2

    ' 2.計算式評価.DK_RNV14
    Dim tmp59(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp59(site) = tmp_DK_RNV13_S2(site) - tmp_DK_RNV14_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV14
    Call ResultAdd("DK_RNV14", tmp59)

' #### DK_RNV15_S2 ####

    ' 0.複数画像情報インポート.DK_RNV15_S2

    ' 1.複数画像情報インポート.DK_RNV15_S2

    ' 2.Subtract(通常).DK_RNV15_S2

    ' 3.ExecuteLUT.DK_RNV15_S2

    ' 8.ShiftLeft.DK_RNV15_S2

    ' 9.複数画像用_LSB定義.DK_RNV15_S2

    ' 12.SliceLevel生成.DK_RNV15_S2
    Dim tmp_Slice16(nSite) As Double
    Call MakeSliceLevel(tmp_Slice16, 0.0016 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV15_S2
    Dim tmp60_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice16, tmp_Slice16, idpLimitEachSite, idpLimitExclude, tmp60_0)
    Dim tmp61 As CImgColorAllResult
    Call GetSum_CImgColor(tmp61, tmp60_0)

    ' 14.GetSum_Color.DK_RNV15_S2
    Dim tmp62(nSite) As Double
    Call GetSum_Color(tmp62, tmp61, "-")

    ' 15.PutTestResult.DK_RNV15_S2
    Call ResultAdd("DK_RNV15_S2", tmp62)

' #### DK_RNV15 ####

    ' 0.測定結果取得.DK_RNV15

    ' 1.測定結果取得.DK_RNV15
    Dim tmp_DK_RNV15_S2() As Double
    TheResult.GetResult "DK_RNV15_S2", tmp_DK_RNV15_S2

    ' 2.計算式評価.DK_RNV15
    Dim tmp63(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp63(site) = tmp_DK_RNV14_S2(site) - tmp_DK_RNV15_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV15
    Call ResultAdd("DK_RNV15", tmp63)

' #### DK_RNV16_S2 ####

    ' 0.複数画像情報インポート.DK_RNV16_S2

    ' 1.複数画像情報インポート.DK_RNV16_S2

    ' 2.Subtract(通常).DK_RNV16_S2

    ' 3.ExecuteLUT.DK_RNV16_S2

    ' 8.ShiftLeft.DK_RNV16_S2

    ' 9.複数画像用_LSB定義.DK_RNV16_S2

    ' 12.SliceLevel生成.DK_RNV16_S2
    Dim tmp_Slice17(nSite) As Double
    Call MakeSliceLevel(tmp_Slice17, 0.0017 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV16_S2
    Dim tmp64_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice17, tmp_Slice17, idpLimitEachSite, idpLimitExclude, tmp64_0)
    Dim tmp65 As CImgColorAllResult
    Call GetSum_CImgColor(tmp65, tmp64_0)

    ' 14.GetSum_Color.DK_RNV16_S2
    Dim tmp66(nSite) As Double
    Call GetSum_Color(tmp66, tmp65, "-")

    ' 15.PutTestResult.DK_RNV16_S2
    Call ResultAdd("DK_RNV16_S2", tmp66)

' #### DK_RNV16 ####

    ' 0.測定結果取得.DK_RNV16

    ' 1.測定結果取得.DK_RNV16
    Dim tmp_DK_RNV16_S2() As Double
    TheResult.GetResult "DK_RNV16_S2", tmp_DK_RNV16_S2

    ' 2.計算式評価.DK_RNV16
    Dim tmp67(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp67(site) = tmp_DK_RNV15_S2(site) - tmp_DK_RNV16_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV16
    Call ResultAdd("DK_RNV16", tmp67)

' #### DK_RNV17_S2 ####

    ' 0.複数画像情報インポート.DK_RNV17_S2

    ' 1.複数画像情報インポート.DK_RNV17_S2

    ' 2.Subtract(通常).DK_RNV17_S2

    ' 3.ExecuteLUT.DK_RNV17_S2

    ' 8.ShiftLeft.DK_RNV17_S2

    ' 9.複数画像用_LSB定義.DK_RNV17_S2

    ' 12.SliceLevel生成.DK_RNV17_S2
    Dim tmp_Slice18(nSite) As Double
    Call MakeSliceLevel(tmp_Slice18, 0.0018 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV17_S2
    Dim tmp68_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice18, tmp_Slice18, idpLimitEachSite, idpLimitExclude, tmp68_0)
    Dim tmp69 As CImgColorAllResult
    Call GetSum_CImgColor(tmp69, tmp68_0)

    ' 14.GetSum_Color.DK_RNV17_S2
    Dim tmp70(nSite) As Double
    Call GetSum_Color(tmp70, tmp69, "-")

    ' 15.PutTestResult.DK_RNV17_S2
    Call ResultAdd("DK_RNV17_S2", tmp70)

' #### DK_RNV17 ####

    ' 0.測定結果取得.DK_RNV17

    ' 1.測定結果取得.DK_RNV17
    Dim tmp_DK_RNV17_S2() As Double
    TheResult.GetResult "DK_RNV17_S2", tmp_DK_RNV17_S2

    ' 2.計算式評価.DK_RNV17
    Dim tmp71(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp71(site) = tmp_DK_RNV16_S2(site) - tmp_DK_RNV17_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV17
    Call ResultAdd("DK_RNV17", tmp71)

' #### DK_RNV18_S2 ####

    ' 0.複数画像情報インポート.DK_RNV18_S2

    ' 1.複数画像情報インポート.DK_RNV18_S2

    ' 2.Subtract(通常).DK_RNV18_S2

    ' 3.ExecuteLUT.DK_RNV18_S2

    ' 8.ShiftLeft.DK_RNV18_S2

    ' 9.複数画像用_LSB定義.DK_RNV18_S2

    ' 12.SliceLevel生成.DK_RNV18_S2
    Dim tmp_Slice19(nSite) As Double
    Call MakeSliceLevel(tmp_Slice19, 0.0019 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV18_S2
    Dim tmp72_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice19, tmp_Slice19, idpLimitEachSite, idpLimitExclude, tmp72_0)
    Dim tmp73 As CImgColorAllResult
    Call GetSum_CImgColor(tmp73, tmp72_0)

    ' 14.GetSum_Color.DK_RNV18_S2
    Dim tmp74(nSite) As Double
    Call GetSum_Color(tmp74, tmp73, "-")

    ' 15.PutTestResult.DK_RNV18_S2
    Call ResultAdd("DK_RNV18_S2", tmp74)

' #### DK_RNV18 ####

    ' 0.測定結果取得.DK_RNV18

    ' 1.測定結果取得.DK_RNV18
    Dim tmp_DK_RNV18_S2() As Double
    TheResult.GetResult "DK_RNV18_S2", tmp_DK_RNV18_S2

    ' 2.計算式評価.DK_RNV18
    Dim tmp75(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp75(site) = tmp_DK_RNV17_S2(site) - tmp_DK_RNV18_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV18
    Call ResultAdd("DK_RNV18", tmp75)

' #### DK_RNV19_S2 ####

    ' 0.複数画像情報インポート.DK_RNV19_S2

    ' 1.複数画像情報インポート.DK_RNV19_S2

    ' 2.Subtract(通常).DK_RNV19_S2

    ' 3.ExecuteLUT.DK_RNV19_S2

    ' 8.ShiftLeft.DK_RNV19_S2

    ' 9.複数画像用_LSB定義.DK_RNV19_S2

    ' 12.SliceLevel生成.DK_RNV19_S2
    Dim tmp_Slice20(nSite) As Double
    Call MakeSliceLevel(tmp_Slice20, 0.002 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV19_S2
    Dim tmp76_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice20, tmp_Slice20, idpLimitEachSite, idpLimitExclude, tmp76_0)
    Dim tmp77 As CImgColorAllResult
    Call GetSum_CImgColor(tmp77, tmp76_0)

    ' 14.GetSum_Color.DK_RNV19_S2
    Dim tmp78(nSite) As Double
    Call GetSum_Color(tmp78, tmp77, "-")

    ' 15.PutTestResult.DK_RNV19_S2
    Call ResultAdd("DK_RNV19_S2", tmp78)

' #### DK_RNV19 ####

    ' 0.測定結果取得.DK_RNV19

    ' 1.測定結果取得.DK_RNV19
    Dim tmp_DK_RNV19_S2() As Double
    TheResult.GetResult "DK_RNV19_S2", tmp_DK_RNV19_S2

    ' 2.計算式評価.DK_RNV19
    Dim tmp79(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp79(site) = tmp_DK_RNV18_S2(site) - tmp_DK_RNV19_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV19
    Call ResultAdd("DK_RNV19", tmp79)

' #### DK_RNV20_S2 ####

    ' 0.複数画像情報インポート.DK_RNV20_S2

    ' 1.複数画像情報インポート.DK_RNV20_S2

    ' 2.Subtract(通常).DK_RNV20_S2

    ' 3.ExecuteLUT.DK_RNV20_S2

    ' 8.ShiftLeft.DK_RNV20_S2

    ' 9.複数画像用_LSB定義.DK_RNV20_S2

    ' 12.SliceLevel生成.DK_RNV20_S2
    Dim tmp_Slice21(nSite) As Double
    Call MakeSliceLevel(tmp_Slice21, 0.0021 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV20_S2
    Dim tmp80_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice21, tmp_Slice21, idpLimitEachSite, idpLimitExclude, tmp80_0)
    Dim tmp81 As CImgColorAllResult
    Call GetSum_CImgColor(tmp81, tmp80_0)

    ' 14.GetSum_Color.DK_RNV20_S2
    Dim tmp82(nSite) As Double
    Call GetSum_Color(tmp82, tmp81, "-")

    ' 15.PutTestResult.DK_RNV20_S2
    Call ResultAdd("DK_RNV20_S2", tmp82)

' #### DK_RNV20 ####

    ' 0.測定結果取得.DK_RNV20

    ' 1.測定結果取得.DK_RNV20
    Dim tmp_DK_RNV20_S2() As Double
    TheResult.GetResult "DK_RNV20_S2", tmp_DK_RNV20_S2

    ' 2.計算式評価.DK_RNV20
    Dim tmp83(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp83(site) = tmp_DK_RNV19_S2(site) - tmp_DK_RNV20_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV20
    Call ResultAdd("DK_RNV20", tmp83)

' #### DK_RNV21_S2 ####

    ' 0.複数画像情報インポート.DK_RNV21_S2

    ' 1.複数画像情報インポート.DK_RNV21_S2

    ' 2.Subtract(通常).DK_RNV21_S2

    ' 3.ExecuteLUT.DK_RNV21_S2

    ' 8.ShiftLeft.DK_RNV21_S2

    ' 9.複数画像用_LSB定義.DK_RNV21_S2

    ' 12.SliceLevel生成.DK_RNV21_S2
    Dim tmp_Slice22(nSite) As Double
    Call MakeSliceLevel(tmp_Slice22, 0.0022 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV21_S2
    Dim tmp84_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice22, tmp_Slice22, idpLimitEachSite, idpLimitExclude, tmp84_0)
    Dim tmp85 As CImgColorAllResult
    Call GetSum_CImgColor(tmp85, tmp84_0)

    ' 14.GetSum_Color.DK_RNV21_S2
    Dim tmp86(nSite) As Double
    Call GetSum_Color(tmp86, tmp85, "-")

    ' 15.PutTestResult.DK_RNV21_S2
    Call ResultAdd("DK_RNV21_S2", tmp86)

' #### DK_RNV21 ####

    ' 0.測定結果取得.DK_RNV21

    ' 1.測定結果取得.DK_RNV21
    Dim tmp_DK_RNV21_S2() As Double
    TheResult.GetResult "DK_RNV21_S2", tmp_DK_RNV21_S2

    ' 2.計算式評価.DK_RNV21
    Dim tmp87(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp87(site) = tmp_DK_RNV20_S2(site) - tmp_DK_RNV21_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV21
    Call ResultAdd("DK_RNV21", tmp87)

' #### DK_RNV22_S2 ####

    ' 0.複数画像情報インポート.DK_RNV22_S2

    ' 1.複数画像情報インポート.DK_RNV22_S2

    ' 2.Subtract(通常).DK_RNV22_S2

    ' 3.ExecuteLUT.DK_RNV22_S2

    ' 8.ShiftLeft.DK_RNV22_S2

    ' 9.複数画像用_LSB定義.DK_RNV22_S2

    ' 12.SliceLevel生成.DK_RNV22_S2
    Dim tmp_Slice23(nSite) As Double
    Call MakeSliceLevel(tmp_Slice23, 0.0023 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV22_S2
    Dim tmp88_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice23, tmp_Slice23, idpLimitEachSite, idpLimitExclude, tmp88_0)
    Dim tmp89 As CImgColorAllResult
    Call GetSum_CImgColor(tmp89, tmp88_0)

    ' 14.GetSum_Color.DK_RNV22_S2
    Dim tmp90(nSite) As Double
    Call GetSum_Color(tmp90, tmp89, "-")

    ' 15.PutTestResult.DK_RNV22_S2
    Call ResultAdd("DK_RNV22_S2", tmp90)

' #### DK_RNV22 ####

    ' 0.測定結果取得.DK_RNV22

    ' 1.測定結果取得.DK_RNV22
    Dim tmp_DK_RNV22_S2() As Double
    TheResult.GetResult "DK_RNV22_S2", tmp_DK_RNV22_S2

    ' 2.計算式評価.DK_RNV22
    Dim tmp91(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp91(site) = tmp_DK_RNV21_S2(site) - tmp_DK_RNV22_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV22
    Call ResultAdd("DK_RNV22", tmp91)

' #### DK_RNV23_S2 ####

    ' 0.複数画像情報インポート.DK_RNV23_S2

    ' 1.複数画像情報インポート.DK_RNV23_S2

    ' 2.Subtract(通常).DK_RNV23_S2

    ' 3.ExecuteLUT.DK_RNV23_S2

    ' 8.ShiftLeft.DK_RNV23_S2

    ' 9.複数画像用_LSB定義.DK_RNV23_S2

    ' 12.SliceLevel生成.DK_RNV23_S2
    Dim tmp_Slice24(nSite) As Double
    Call MakeSliceLevel(tmp_Slice24, 0.0024 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV23_S2
    Dim tmp92_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice24, tmp_Slice24, idpLimitEachSite, idpLimitExclude, tmp92_0)
    Dim tmp93 As CImgColorAllResult
    Call GetSum_CImgColor(tmp93, tmp92_0)

    ' 14.GetSum_Color.DK_RNV23_S2
    Dim tmp94(nSite) As Double
    Call GetSum_Color(tmp94, tmp93, "-")

    ' 15.PutTestResult.DK_RNV23_S2
    Call ResultAdd("DK_RNV23_S2", tmp94)

' #### DK_RNV23 ####

    ' 0.測定結果取得.DK_RNV23

    ' 1.測定結果取得.DK_RNV23
    Dim tmp_DK_RNV23_S2() As Double
    TheResult.GetResult "DK_RNV23_S2", tmp_DK_RNV23_S2

    ' 2.計算式評価.DK_RNV23
    Dim tmp95(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp95(site) = tmp_DK_RNV22_S2(site) - tmp_DK_RNV23_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV23
    Call ResultAdd("DK_RNV23", tmp95)

' #### DK_RNV24_S2 ####

    ' 0.複数画像情報インポート.DK_RNV24_S2

    ' 1.複数画像情報インポート.DK_RNV24_S2

    ' 2.Subtract(通常).DK_RNV24_S2

    ' 3.ExecuteLUT.DK_RNV24_S2

    ' 8.ShiftLeft.DK_RNV24_S2

    ' 9.複数画像用_LSB定義.DK_RNV24_S2

    ' 12.SliceLevel生成.DK_RNV24_S2
    Dim tmp_Slice25(nSite) As Double
    Call MakeSliceLevel(tmp_Slice25, 0.0025 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV24_S2
    Dim tmp96_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice25, tmp_Slice25, idpLimitEachSite, idpLimitExclude, tmp96_0)
    Dim tmp97 As CImgColorAllResult
    Call GetSum_CImgColor(tmp97, tmp96_0)

    ' 14.GetSum_Color.DK_RNV24_S2
    Dim tmp98(nSite) As Double
    Call GetSum_Color(tmp98, tmp97, "-")

    ' 15.PutTestResult.DK_RNV24_S2
    Call ResultAdd("DK_RNV24_S2", tmp98)

' #### DK_RNV24 ####

    ' 0.測定結果取得.DK_RNV24

    ' 1.測定結果取得.DK_RNV24
    Dim tmp_DK_RNV24_S2() As Double
    TheResult.GetResult "DK_RNV24_S2", tmp_DK_RNV24_S2

    ' 2.計算式評価.DK_RNV24
    Dim tmp99(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp99(site) = tmp_DK_RNV23_S2(site) - tmp_DK_RNV24_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV24
    Call ResultAdd("DK_RNV24", tmp99)

' #### DK_RNV25_S2 ####

    ' 0.複数画像情報インポート.DK_RNV25_S2

    ' 1.複数画像情報インポート.DK_RNV25_S2

    ' 2.Subtract(通常).DK_RNV25_S2

    ' 3.ExecuteLUT.DK_RNV25_S2

    ' 8.ShiftLeft.DK_RNV25_S2

    ' 9.複数画像用_LSB定義.DK_RNV25_S2

    ' 12.SliceLevel生成.DK_RNV25_S2
    Dim tmp_Slice26(nSite) As Double
    Call MakeSliceLevel(tmp_Slice26, 0.0026 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV25_S2
    Dim tmp100_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice26, tmp_Slice26, idpLimitEachSite, idpLimitExclude, tmp100_0)
    Dim tmp101 As CImgColorAllResult
    Call GetSum_CImgColor(tmp101, tmp100_0)

    ' 14.GetSum_Color.DK_RNV25_S2
    Dim tmp102(nSite) As Double
    Call GetSum_Color(tmp102, tmp101, "-")

    ' 15.PutTestResult.DK_RNV25_S2
    Call ResultAdd("DK_RNV25_S2", tmp102)

' #### DK_RNV25 ####

    ' 0.測定結果取得.DK_RNV25

    ' 1.測定結果取得.DK_RNV25
    Dim tmp_DK_RNV25_S2() As Double
    TheResult.GetResult "DK_RNV25_S2", tmp_DK_RNV25_S2

    ' 2.計算式評価.DK_RNV25
    Dim tmp103(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp103(site) = tmp_DK_RNV24_S2(site) - tmp_DK_RNV25_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV25
    Call ResultAdd("DK_RNV25", tmp103)

' #### DK_RNV26_S2 ####

    ' 0.複数画像情報インポート.DK_RNV26_S2

    ' 1.複数画像情報インポート.DK_RNV26_S2

    ' 2.Subtract(通常).DK_RNV26_S2

    ' 3.ExecuteLUT.DK_RNV26_S2

    ' 8.ShiftLeft.DK_RNV26_S2

    ' 9.複数画像用_LSB定義.DK_RNV26_S2

    ' 12.SliceLevel生成.DK_RNV26_S2
    Dim tmp_Slice27(nSite) As Double
    Call MakeSliceLevel(tmp_Slice27, 0.0027 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV26_S2
    Dim tmp104_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice27, tmp_Slice27, idpLimitEachSite, idpLimitExclude, tmp104_0)
    Dim tmp105 As CImgColorAllResult
    Call GetSum_CImgColor(tmp105, tmp104_0)

    ' 14.GetSum_Color.DK_RNV26_S2
    Dim tmp106(nSite) As Double
    Call GetSum_Color(tmp106, tmp105, "-")

    ' 15.PutTestResult.DK_RNV26_S2
    Call ResultAdd("DK_RNV26_S2", tmp106)

' #### DK_RNV26 ####

    ' 0.測定結果取得.DK_RNV26

    ' 1.測定結果取得.DK_RNV26
    Dim tmp_DK_RNV26_S2() As Double
    TheResult.GetResult "DK_RNV26_S2", tmp_DK_RNV26_S2

    ' 2.計算式評価.DK_RNV26
    Dim tmp107(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp107(site) = tmp_DK_RNV25_S2(site) - tmp_DK_RNV26_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV26
    Call ResultAdd("DK_RNV26", tmp107)

' #### DK_RNV27_S2 ####

    ' 0.複数画像情報インポート.DK_RNV27_S2

    ' 1.複数画像情報インポート.DK_RNV27_S2

    ' 2.Subtract(通常).DK_RNV27_S2

    ' 3.ExecuteLUT.DK_RNV27_S2

    ' 8.ShiftLeft.DK_RNV27_S2

    ' 9.複数画像用_LSB定義.DK_RNV27_S2

    ' 12.SliceLevel生成.DK_RNV27_S2
    Dim tmp_Slice28(nSite) As Double
    Call MakeSliceLevel(tmp_Slice28, 0.0028 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV27_S2
    Dim tmp108_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice28, tmp_Slice28, idpLimitEachSite, idpLimitExclude, tmp108_0)
    Dim tmp109 As CImgColorAllResult
    Call GetSum_CImgColor(tmp109, tmp108_0)

    ' 14.GetSum_Color.DK_RNV27_S2
    Dim tmp110(nSite) As Double
    Call GetSum_Color(tmp110, tmp109, "-")

    ' 15.PutTestResult.DK_RNV27_S2
    Call ResultAdd("DK_RNV27_S2", tmp110)

' #### DK_RNV27 ####

    ' 0.測定結果取得.DK_RNV27

    ' 1.測定結果取得.DK_RNV27
    Dim tmp_DK_RNV27_S2() As Double
    TheResult.GetResult "DK_RNV27_S2", tmp_DK_RNV27_S2

    ' 2.計算式評価.DK_RNV27
    Dim tmp111(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp111(site) = tmp_DK_RNV26_S2(site) - tmp_DK_RNV27_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV27
    Call ResultAdd("DK_RNV27", tmp111)

' #### DK_RNV28_S2 ####

    ' 0.複数画像情報インポート.DK_RNV28_S2

    ' 1.複数画像情報インポート.DK_RNV28_S2

    ' 2.Subtract(通常).DK_RNV28_S2

    ' 3.ExecuteLUT.DK_RNV28_S2

    ' 8.ShiftLeft.DK_RNV28_S2

    ' 9.複数画像用_LSB定義.DK_RNV28_S2

    ' 12.SliceLevel生成.DK_RNV28_S2
    Dim tmp_Slice29(nSite) As Double
    Call MakeSliceLevel(tmp_Slice29, 0.0029 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV28_S2
    Dim tmp112_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice29, tmp_Slice29, idpLimitEachSite, idpLimitExclude, tmp112_0)
    Dim tmp113 As CImgColorAllResult
    Call GetSum_CImgColor(tmp113, tmp112_0)

    ' 14.GetSum_Color.DK_RNV28_S2
    Dim tmp114(nSite) As Double
    Call GetSum_Color(tmp114, tmp113, "-")

    ' 15.PutTestResult.DK_RNV28_S2
    Call ResultAdd("DK_RNV28_S2", tmp114)

' #### DK_RNV28 ####

    ' 0.測定結果取得.DK_RNV28

    ' 1.測定結果取得.DK_RNV28
    Dim tmp_DK_RNV28_S2() As Double
    TheResult.GetResult "DK_RNV28_S2", tmp_DK_RNV28_S2

    ' 2.計算式評価.DK_RNV28
    Dim tmp115(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp115(site) = tmp_DK_RNV27_S2(site) - tmp_DK_RNV28_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV28
    Call ResultAdd("DK_RNV28", tmp115)

' #### DK_RNV29_S2 ####

    ' 0.複数画像情報インポート.DK_RNV29_S2

    ' 1.複数画像情報インポート.DK_RNV29_S2

    ' 2.Subtract(通常).DK_RNV29_S2

    ' 3.ExecuteLUT.DK_RNV29_S2

    ' 8.ShiftLeft.DK_RNV29_S2

    ' 9.複数画像用_LSB定義.DK_RNV29_S2

    ' 12.SliceLevel生成.DK_RNV29_S2
    Dim tmp_Slice30(nSite) As Double
    Call MakeSliceLevel(tmp_Slice30, 0.003 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV29_S2
    Dim tmp116_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice30, tmp_Slice30, idpLimitEachSite, idpLimitExclude, tmp116_0)
    Dim tmp117 As CImgColorAllResult
    Call GetSum_CImgColor(tmp117, tmp116_0)

    ' 14.GetSum_Color.DK_RNV29_S2
    Dim tmp118(nSite) As Double
    Call GetSum_Color(tmp118, tmp117, "-")

    ' 15.PutTestResult.DK_RNV29_S2
    Call ResultAdd("DK_RNV29_S2", tmp118)

' #### DK_RNV29 ####

    ' 0.測定結果取得.DK_RNV29

    ' 1.測定結果取得.DK_RNV29
    Dim tmp_DK_RNV29_S2() As Double
    TheResult.GetResult "DK_RNV29_S2", tmp_DK_RNV29_S2

    ' 2.計算式評価.DK_RNV29
    Dim tmp119(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp119(site) = tmp_DK_RNV28_S2(site) - tmp_DK_RNV29_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV29
    Call ResultAdd("DK_RNV29", tmp119)

' #### DK_RNV30_S2 ####

    ' 0.複数画像情報インポート.DK_RNV30_S2

    ' 1.複数画像情報インポート.DK_RNV30_S2

    ' 2.Subtract(通常).DK_RNV30_S2

    ' 3.ExecuteLUT.DK_RNV30_S2

    ' 8.ShiftLeft.DK_RNV30_S2

    ' 9.複数画像用_LSB定義.DK_RNV30_S2

    ' 12.SliceLevel生成.DK_RNV30_S2
    Dim tmp_Slice31(nSite) As Double
    Call MakeSliceLevel(tmp_Slice31, 0.0031 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV30_S2
    Dim tmp120_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice31, tmp_Slice31, idpLimitEachSite, idpLimitExclude, tmp120_0)
    Dim tmp121 As CImgColorAllResult
    Call GetSum_CImgColor(tmp121, tmp120_0)

    ' 14.GetSum_Color.DK_RNV30_S2
    Dim tmp122(nSite) As Double
    Call GetSum_Color(tmp122, tmp121, "-")

    ' 15.PutTestResult.DK_RNV30_S2
    Call ResultAdd("DK_RNV30_S2", tmp122)

' #### DK_RNV30 ####

    ' 0.測定結果取得.DK_RNV30

    ' 1.測定結果取得.DK_RNV30
    Dim tmp_DK_RNV30_S2() As Double
    TheResult.GetResult "DK_RNV30_S2", tmp_DK_RNV30_S2

    ' 2.計算式評価.DK_RNV30
    Dim tmp123(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp123(site) = tmp_DK_RNV29_S2(site) - tmp_DK_RNV30_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV30
    Call ResultAdd("DK_RNV30", tmp123)

' #### DK_RNV31_S2 ####

    ' 0.複数画像情報インポート.DK_RNV31_S2

    ' 1.複数画像情報インポート.DK_RNV31_S2

    ' 2.Subtract(通常).DK_RNV31_S2

    ' 3.ExecuteLUT.DK_RNV31_S2

    ' 8.ShiftLeft.DK_RNV31_S2

    ' 9.複数画像用_LSB定義.DK_RNV31_S2

    ' 12.SliceLevel生成.DK_RNV31_S2
    Dim tmp_Slice32(nSite) As Double
    Call MakeSliceLevel(tmp_Slice32, 0.0032 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV31_S2
    Dim tmp124_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice32, tmp_Slice32, idpLimitEachSite, idpLimitExclude, tmp124_0)
    Dim tmp125 As CImgColorAllResult
    Call GetSum_CImgColor(tmp125, tmp124_0)

    ' 14.GetSum_Color.DK_RNV31_S2
    Dim tmp126(nSite) As Double
    Call GetSum_Color(tmp126, tmp125, "-")

    ' 15.PutTestResult.DK_RNV31_S2
    Call ResultAdd("DK_RNV31_S2", tmp126)

' #### DK_RNV31 ####

    ' 0.測定結果取得.DK_RNV31

    ' 1.測定結果取得.DK_RNV31
    Dim tmp_DK_RNV31_S2() As Double
    TheResult.GetResult "DK_RNV31_S2", tmp_DK_RNV31_S2

    ' 2.計算式評価.DK_RNV31
    Dim tmp127(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp127(site) = tmp_DK_RNV30_S2(site) - tmp_DK_RNV31_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV31
    Call ResultAdd("DK_RNV31", tmp127)

' #### DK_RNV32_S2 ####

    ' 0.複数画像情報インポート.DK_RNV32_S2

    ' 1.複数画像情報インポート.DK_RNV32_S2

    ' 2.Subtract(通常).DK_RNV32_S2

    ' 3.ExecuteLUT.DK_RNV32_S2

    ' 8.ShiftLeft.DK_RNV32_S2

    ' 9.複数画像用_LSB定義.DK_RNV32_S2

    ' 12.SliceLevel生成.DK_RNV32_S2
    Dim tmp_Slice33(nSite) As Double
    Call MakeSliceLevel(tmp_Slice33, 0.0033 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV32_S2
    Dim tmp128_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice33, tmp_Slice33, idpLimitEachSite, idpLimitExclude, tmp128_0)
    Dim tmp129 As CImgColorAllResult
    Call GetSum_CImgColor(tmp129, tmp128_0)

    ' 14.GetSum_Color.DK_RNV32_S2
    Dim tmp130(nSite) As Double
    Call GetSum_Color(tmp130, tmp129, "-")

    ' 15.PutTestResult.DK_RNV32_S2
    Call ResultAdd("DK_RNV32_S2", tmp130)

' #### DK_RNV32 ####

    ' 0.測定結果取得.DK_RNV32

    ' 1.測定結果取得.DK_RNV32
    Dim tmp_DK_RNV32_S2() As Double
    TheResult.GetResult "DK_RNV32_S2", tmp_DK_RNV32_S2

    ' 2.計算式評価.DK_RNV32
    Dim tmp131(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp131(site) = tmp_DK_RNV31_S2(site) - tmp_DK_RNV32_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV32
    Call ResultAdd("DK_RNV32", tmp131)

' #### DK_RNV33_S2 ####

    ' 0.複数画像情報インポート.DK_RNV33_S2

    ' 1.複数画像情報インポート.DK_RNV33_S2

    ' 2.Subtract(通常).DK_RNV33_S2

    ' 3.ExecuteLUT.DK_RNV33_S2

    ' 8.ShiftLeft.DK_RNV33_S2

    ' 9.複数画像用_LSB定義.DK_RNV33_S2

    ' 12.SliceLevel生成.DK_RNV33_S2
    Dim tmp_Slice34(nSite) As Double
    Call MakeSliceLevel(tmp_Slice34, 0.0034 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV33_S2
    Dim tmp132_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice34, tmp_Slice34, idpLimitEachSite, idpLimitExclude, tmp132_0)
    Dim tmp133 As CImgColorAllResult
    Call GetSum_CImgColor(tmp133, tmp132_0)

    ' 14.GetSum_Color.DK_RNV33_S2
    Dim tmp134(nSite) As Double
    Call GetSum_Color(tmp134, tmp133, "-")

    ' 15.PutTestResult.DK_RNV33_S2
    Call ResultAdd("DK_RNV33_S2", tmp134)

' #### DK_RNV33 ####

    ' 0.測定結果取得.DK_RNV33

    ' 1.測定結果取得.DK_RNV33
    Dim tmp_DK_RNV33_S2() As Double
    TheResult.GetResult "DK_RNV33_S2", tmp_DK_RNV33_S2

    ' 2.計算式評価.DK_RNV33
    Dim tmp135(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp135(site) = tmp_DK_RNV32_S2(site) - tmp_DK_RNV33_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV33
    Call ResultAdd("DK_RNV33", tmp135)

' #### DK_RNV34_S2 ####

    ' 0.複数画像情報インポート.DK_RNV34_S2

    ' 1.複数画像情報インポート.DK_RNV34_S2

    ' 2.Subtract(通常).DK_RNV34_S2

    ' 3.ExecuteLUT.DK_RNV34_S2

    ' 8.ShiftLeft.DK_RNV34_S2

    ' 9.複数画像用_LSB定義.DK_RNV34_S2

    ' 12.SliceLevel生成.DK_RNV34_S2
    Dim tmp_Slice35(nSite) As Double
    Call MakeSliceLevel(tmp_Slice35, 0.0035 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV34_S2
    Dim tmp136_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice35, tmp_Slice35, idpLimitEachSite, idpLimitExclude, tmp136_0)
    Dim tmp137 As CImgColorAllResult
    Call GetSum_CImgColor(tmp137, tmp136_0)

    ' 14.GetSum_Color.DK_RNV34_S2
    Dim tmp138(nSite) As Double
    Call GetSum_Color(tmp138, tmp137, "-")

    ' 15.PutTestResult.DK_RNV34_S2
    Call ResultAdd("DK_RNV34_S2", tmp138)

' #### DK_RNV34 ####

    ' 0.測定結果取得.DK_RNV34

    ' 1.測定結果取得.DK_RNV34
    Dim tmp_DK_RNV34_S2() As Double
    TheResult.GetResult "DK_RNV34_S2", tmp_DK_RNV34_S2

    ' 2.計算式評価.DK_RNV34
    Dim tmp139(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp139(site) = tmp_DK_RNV33_S2(site) - tmp_DK_RNV34_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV34
    Call ResultAdd("DK_RNV34", tmp139)

' #### DK_RNV35_S2 ####

    ' 0.複数画像情報インポート.DK_RNV35_S2

    ' 1.複数画像情報インポート.DK_RNV35_S2

    ' 2.Subtract(通常).DK_RNV35_S2

    ' 3.ExecuteLUT.DK_RNV35_S2

    ' 8.ShiftLeft.DK_RNV35_S2

    ' 9.複数画像用_LSB定義.DK_RNV35_S2

    ' 12.SliceLevel生成.DK_RNV35_S2
    Dim tmp_Slice36(nSite) As Double
    Call MakeSliceLevel(tmp_Slice36, 0.0036 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV35_S2
    Dim tmp140_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice36, tmp_Slice36, idpLimitEachSite, idpLimitExclude, tmp140_0)
    Dim tmp141 As CImgColorAllResult
    Call GetSum_CImgColor(tmp141, tmp140_0)

    ' 14.GetSum_Color.DK_RNV35_S2
    Dim tmp142(nSite) As Double
    Call GetSum_Color(tmp142, tmp141, "-")

    ' 15.PutTestResult.DK_RNV35_S2
    Call ResultAdd("DK_RNV35_S2", tmp142)

' #### DK_RNV35 ####

    ' 0.測定結果取得.DK_RNV35

    ' 1.測定結果取得.DK_RNV35
    Dim tmp_DK_RNV35_S2() As Double
    TheResult.GetResult "DK_RNV35_S2", tmp_DK_RNV35_S2

    ' 2.計算式評価.DK_RNV35
    Dim tmp143(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp143(site) = tmp_DK_RNV34_S2(site) - tmp_DK_RNV35_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV35
    Call ResultAdd("DK_RNV35", tmp143)

' #### DK_RNV36_S2 ####

    ' 0.複数画像情報インポート.DK_RNV36_S2

    ' 1.複数画像情報インポート.DK_RNV36_S2

    ' 2.Subtract(通常).DK_RNV36_S2

    ' 3.ExecuteLUT.DK_RNV36_S2

    ' 8.ShiftLeft.DK_RNV36_S2

    ' 9.複数画像用_LSB定義.DK_RNV36_S2

    ' 12.SliceLevel生成.DK_RNV36_S2
    Dim tmp_Slice37(nSite) As Double
    Call MakeSliceLevel(tmp_Slice37, 0.0037 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV36_S2
    Dim tmp144_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice37, tmp_Slice37, idpLimitEachSite, idpLimitExclude, tmp144_0)
    Dim tmp145 As CImgColorAllResult
    Call GetSum_CImgColor(tmp145, tmp144_0)

    ' 14.GetSum_Color.DK_RNV36_S2
    Dim tmp146(nSite) As Double
    Call GetSum_Color(tmp146, tmp145, "-")

    ' 15.PutTestResult.DK_RNV36_S2
    Call ResultAdd("DK_RNV36_S2", tmp146)

' #### DK_RNV36 ####

    ' 0.測定結果取得.DK_RNV36

    ' 1.測定結果取得.DK_RNV36
    Dim tmp_DK_RNV36_S2() As Double
    TheResult.GetResult "DK_RNV36_S2", tmp_DK_RNV36_S2

    ' 2.計算式評価.DK_RNV36
    Dim tmp147(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp147(site) = tmp_DK_RNV35_S2(site) - tmp_DK_RNV36_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV36
    Call ResultAdd("DK_RNV36", tmp147)

' #### DK_RNV37_S2 ####

    ' 0.複数画像情報インポート.DK_RNV37_S2

    ' 1.複数画像情報インポート.DK_RNV37_S2

    ' 2.Subtract(通常).DK_RNV37_S2

    ' 3.ExecuteLUT.DK_RNV37_S2

    ' 8.ShiftLeft.DK_RNV37_S2

    ' 9.複数画像用_LSB定義.DK_RNV37_S2

    ' 12.SliceLevel生成.DK_RNV37_S2
    Dim tmp_Slice38(nSite) As Double
    Call MakeSliceLevel(tmp_Slice38, 0.0038 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV37_S2
    Dim tmp148_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice38, tmp_Slice38, idpLimitEachSite, idpLimitExclude, tmp148_0)
    Dim tmp149 As CImgColorAllResult
    Call GetSum_CImgColor(tmp149, tmp148_0)

    ' 14.GetSum_Color.DK_RNV37_S2
    Dim tmp150(nSite) As Double
    Call GetSum_Color(tmp150, tmp149, "-")

    ' 15.PutTestResult.DK_RNV37_S2
    Call ResultAdd("DK_RNV37_S2", tmp150)

' #### DK_RNV37 ####

    ' 0.測定結果取得.DK_RNV37

    ' 1.測定結果取得.DK_RNV37
    Dim tmp_DK_RNV37_S2() As Double
    TheResult.GetResult "DK_RNV37_S2", tmp_DK_RNV37_S2

    ' 2.計算式評価.DK_RNV37
    Dim tmp151(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp151(site) = tmp_DK_RNV36_S2(site) - tmp_DK_RNV37_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV37
    Call ResultAdd("DK_RNV37", tmp151)

' #### DK_RNV38_S2 ####

    ' 0.複数画像情報インポート.DK_RNV38_S2

    ' 1.複数画像情報インポート.DK_RNV38_S2

    ' 2.Subtract(通常).DK_RNV38_S2

    ' 3.ExecuteLUT.DK_RNV38_S2

    ' 8.ShiftLeft.DK_RNV38_S2

    ' 9.複数画像用_LSB定義.DK_RNV38_S2

    ' 12.SliceLevel生成.DK_RNV38_S2
    Dim tmp_Slice39(nSite) As Double
    Call MakeSliceLevel(tmp_Slice39, 0.0039 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV38_S2
    Dim tmp152_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice39, tmp_Slice39, idpLimitEachSite, idpLimitExclude, tmp152_0)
    Dim tmp153 As CImgColorAllResult
    Call GetSum_CImgColor(tmp153, tmp152_0)

    ' 14.GetSum_Color.DK_RNV38_S2
    Dim tmp154(nSite) As Double
    Call GetSum_Color(tmp154, tmp153, "-")

    ' 15.PutTestResult.DK_RNV38_S2
    Call ResultAdd("DK_RNV38_S2", tmp154)

' #### DK_RNV38 ####

    ' 0.測定結果取得.DK_RNV38

    ' 1.測定結果取得.DK_RNV38
    Dim tmp_DK_RNV38_S2() As Double
    TheResult.GetResult "DK_RNV38_S2", tmp_DK_RNV38_S2

    ' 2.計算式評価.DK_RNV38
    Dim tmp155(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp155(site) = tmp_DK_RNV37_S2(site) - tmp_DK_RNV38_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV38
    Call ResultAdd("DK_RNV38", tmp155)

' #### DK_RNV39_S2 ####

    ' 0.複数画像情報インポート.DK_RNV39_S2

    ' 1.複数画像情報インポート.DK_RNV39_S2

    ' 2.Subtract(通常).DK_RNV39_S2

    ' 3.ExecuteLUT.DK_RNV39_S2

    ' 8.ShiftLeft.DK_RNV39_S2

    ' 9.複数画像用_LSB定義.DK_RNV39_S2

    ' 12.SliceLevel生成.DK_RNV39_S2
    Dim tmp_Slice40(nSite) As Double
    Call MakeSliceLevel(tmp_Slice40, 0.004 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV39_S2
    Dim tmp156_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice40, tmp_Slice40, idpLimitEachSite, idpLimitExclude, tmp156_0)
    Dim tmp157 As CImgColorAllResult
    Call GetSum_CImgColor(tmp157, tmp156_0)

    ' 14.GetSum_Color.DK_RNV39_S2
    Dim tmp158(nSite) As Double
    Call GetSum_Color(tmp158, tmp157, "-")

    ' 15.PutTestResult.DK_RNV39_S2
    Call ResultAdd("DK_RNV39_S2", tmp158)

' #### DK_RNV39 ####

    ' 0.測定結果取得.DK_RNV39

    ' 1.測定結果取得.DK_RNV39
    Dim tmp_DK_RNV39_S2() As Double
    TheResult.GetResult "DK_RNV39_S2", tmp_DK_RNV39_S2

    ' 2.計算式評価.DK_RNV39
    Dim tmp159(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp159(site) = tmp_DK_RNV38_S2(site) - tmp_DK_RNV39_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV39
    Call ResultAdd("DK_RNV39", tmp159)

' #### DK_RNV40_S2 ####

    ' 0.複数画像情報インポート.DK_RNV40_S2

    ' 1.複数画像情報インポート.DK_RNV40_S2

    ' 2.Subtract(通常).DK_RNV40_S2

    ' 3.ExecuteLUT.DK_RNV40_S2

    ' 8.ShiftLeft.DK_RNV40_S2

    ' 9.複数画像用_LSB定義.DK_RNV40_S2

    ' 12.SliceLevel生成.DK_RNV40_S2
    Dim tmp_Slice41(nSite) As Double
    Call MakeSliceLevel(tmp_Slice41, 0.0041 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV40_S2
    Dim tmp160_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice41, tmp_Slice41, idpLimitEachSite, idpLimitExclude, tmp160_0)
    Dim tmp161 As CImgColorAllResult
    Call GetSum_CImgColor(tmp161, tmp160_0)

    ' 14.GetSum_Color.DK_RNV40_S2
    Dim tmp162(nSite) As Double
    Call GetSum_Color(tmp162, tmp161, "-")

    ' 15.PutTestResult.DK_RNV40_S2
    Call ResultAdd("DK_RNV40_S2", tmp162)

' #### DK_RNV40 ####

    ' 0.測定結果取得.DK_RNV40

    ' 1.測定結果取得.DK_RNV40
    Dim tmp_DK_RNV40_S2() As Double
    TheResult.GetResult "DK_RNV40_S2", tmp_DK_RNV40_S2

    ' 2.計算式評価.DK_RNV40
    Dim tmp163(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp163(site) = tmp_DK_RNV39_S2(site) - tmp_DK_RNV40_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV40
    Call ResultAdd("DK_RNV40", tmp163)

' #### DK_RNV41_S2 ####

    ' 0.複数画像情報インポート.DK_RNV41_S2

    ' 1.複数画像情報インポート.DK_RNV41_S2

    ' 2.Subtract(通常).DK_RNV41_S2

    ' 3.ExecuteLUT.DK_RNV41_S2

    ' 8.ShiftLeft.DK_RNV41_S2

    ' 9.複数画像用_LSB定義.DK_RNV41_S2

    ' 12.SliceLevel生成.DK_RNV41_S2
    Dim tmp_Slice42(nSite) As Double
    Call MakeSliceLevel(tmp_Slice42, 0.0042 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV41_S2
    Dim tmp164_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice42, tmp_Slice42, idpLimitEachSite, idpLimitExclude, tmp164_0)
    Dim tmp165 As CImgColorAllResult
    Call GetSum_CImgColor(tmp165, tmp164_0)

    ' 14.GetSum_Color.DK_RNV41_S2
    Dim tmp166(nSite) As Double
    Call GetSum_Color(tmp166, tmp165, "-")

    ' 15.PutTestResult.DK_RNV41_S2
    Call ResultAdd("DK_RNV41_S2", tmp166)

' #### DK_RNV41 ####

    ' 0.測定結果取得.DK_RNV41

    ' 1.測定結果取得.DK_RNV41
    Dim tmp_DK_RNV41_S2() As Double
    TheResult.GetResult "DK_RNV41_S2", tmp_DK_RNV41_S2

    ' 2.計算式評価.DK_RNV41
    Dim tmp167(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp167(site) = tmp_DK_RNV40_S2(site) - tmp_DK_RNV41_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV41
    Call ResultAdd("DK_RNV41", tmp167)

' #### DK_RNV42_S2 ####

    ' 0.複数画像情報インポート.DK_RNV42_S2

    ' 1.複数画像情報インポート.DK_RNV42_S2

    ' 2.Subtract(通常).DK_RNV42_S2

    ' 3.ExecuteLUT.DK_RNV42_S2

    ' 8.ShiftLeft.DK_RNV42_S2

    ' 9.複数画像用_LSB定義.DK_RNV42_S2

    ' 12.SliceLevel生成.DK_RNV42_S2
    Dim tmp_Slice43(nSite) As Double
    Call MakeSliceLevel(tmp_Slice43, 0.0043 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV42_S2
    Dim tmp168_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice43, tmp_Slice43, idpLimitEachSite, idpLimitExclude, tmp168_0)
    Dim tmp169 As CImgColorAllResult
    Call GetSum_CImgColor(tmp169, tmp168_0)

    ' 14.GetSum_Color.DK_RNV42_S2
    Dim tmp170(nSite) As Double
    Call GetSum_Color(tmp170, tmp169, "-")

    ' 15.PutTestResult.DK_RNV42_S2
    Call ResultAdd("DK_RNV42_S2", tmp170)

' #### DK_RNV42 ####

    ' 0.測定結果取得.DK_RNV42

    ' 1.測定結果取得.DK_RNV42
    Dim tmp_DK_RNV42_S2() As Double
    TheResult.GetResult "DK_RNV42_S2", tmp_DK_RNV42_S2

    ' 2.計算式評価.DK_RNV42
    Dim tmp171(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp171(site) = tmp_DK_RNV41_S2(site) - tmp_DK_RNV42_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV42
    Call ResultAdd("DK_RNV42", tmp171)

' #### DK_RNV43_S2 ####

    ' 0.複数画像情報インポート.DK_RNV43_S2

    ' 1.複数画像情報インポート.DK_RNV43_S2

    ' 2.Subtract(通常).DK_RNV43_S2

    ' 3.ExecuteLUT.DK_RNV43_S2

    ' 8.ShiftLeft.DK_RNV43_S2

    ' 9.複数画像用_LSB定義.DK_RNV43_S2

    ' 12.SliceLevel生成.DK_RNV43_S2
    Dim tmp_Slice44(nSite) As Double
    Call MakeSliceLevel(tmp_Slice44, 0.0044 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV43_S2
    Dim tmp172_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice44, tmp_Slice44, idpLimitEachSite, idpLimitExclude, tmp172_0)
    Dim tmp173 As CImgColorAllResult
    Call GetSum_CImgColor(tmp173, tmp172_0)

    ' 14.GetSum_Color.DK_RNV43_S2
    Dim tmp174(nSite) As Double
    Call GetSum_Color(tmp174, tmp173, "-")

    ' 15.PutTestResult.DK_RNV43_S2
    Call ResultAdd("DK_RNV43_S2", tmp174)

' #### DK_RNV43 ####

    ' 0.測定結果取得.DK_RNV43

    ' 1.測定結果取得.DK_RNV43
    Dim tmp_DK_RNV43_S2() As Double
    TheResult.GetResult "DK_RNV43_S2", tmp_DK_RNV43_S2

    ' 2.計算式評価.DK_RNV43
    Dim tmp175(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp175(site) = tmp_DK_RNV42_S2(site) - tmp_DK_RNV43_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV43
    Call ResultAdd("DK_RNV43", tmp175)

' #### DK_RNV44_S2 ####

    ' 0.複数画像情報インポート.DK_RNV44_S2

    ' 1.複数画像情報インポート.DK_RNV44_S2

    ' 2.Subtract(通常).DK_RNV44_S2

    ' 3.ExecuteLUT.DK_RNV44_S2

    ' 8.ShiftLeft.DK_RNV44_S2

    ' 9.複数画像用_LSB定義.DK_RNV44_S2

    ' 12.SliceLevel生成.DK_RNV44_S2
    Dim tmp_Slice45(nSite) As Double
    Call MakeSliceLevel(tmp_Slice45, 0.0045 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV44_S2
    Dim tmp176_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice45, tmp_Slice45, idpLimitEachSite, idpLimitExclude, tmp176_0)
    Dim tmp177 As CImgColorAllResult
    Call GetSum_CImgColor(tmp177, tmp176_0)

    ' 14.GetSum_Color.DK_RNV44_S2
    Dim tmp178(nSite) As Double
    Call GetSum_Color(tmp178, tmp177, "-")

    ' 15.PutTestResult.DK_RNV44_S2
    Call ResultAdd("DK_RNV44_S2", tmp178)

' #### DK_RNV44 ####

    ' 0.測定結果取得.DK_RNV44

    ' 1.測定結果取得.DK_RNV44
    Dim tmp_DK_RNV44_S2() As Double
    TheResult.GetResult "DK_RNV44_S2", tmp_DK_RNV44_S2

    ' 2.計算式評価.DK_RNV44
    Dim tmp179(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp179(site) = tmp_DK_RNV43_S2(site) - tmp_DK_RNV44_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV44
    Call ResultAdd("DK_RNV44", tmp179)

' #### DK_RNV45_S2 ####

    ' 0.複数画像情報インポート.DK_RNV45_S2

    ' 1.複数画像情報インポート.DK_RNV45_S2

    ' 2.Subtract(通常).DK_RNV45_S2

    ' 3.ExecuteLUT.DK_RNV45_S2

    ' 8.ShiftLeft.DK_RNV45_S2

    ' 9.複数画像用_LSB定義.DK_RNV45_S2

    ' 12.SliceLevel生成.DK_RNV45_S2
    Dim tmp_Slice46(nSite) As Double
    Call MakeSliceLevel(tmp_Slice46, 0.0046 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV45_S2
    Dim tmp180_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice46, tmp_Slice46, idpLimitEachSite, idpLimitExclude, tmp180_0)
    Dim tmp181 As CImgColorAllResult
    Call GetSum_CImgColor(tmp181, tmp180_0)

    ' 14.GetSum_Color.DK_RNV45_S2
    Dim tmp182(nSite) As Double
    Call GetSum_Color(tmp182, tmp181, "-")

    ' 15.PutTestResult.DK_RNV45_S2
    Call ResultAdd("DK_RNV45_S2", tmp182)

' #### DK_RNV45 ####

    ' 0.測定結果取得.DK_RNV45

    ' 1.測定結果取得.DK_RNV45
    Dim tmp_DK_RNV45_S2() As Double
    TheResult.GetResult "DK_RNV45_S2", tmp_DK_RNV45_S2

    ' 2.計算式評価.DK_RNV45
    Dim tmp183(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp183(site) = tmp_DK_RNV44_S2(site) - tmp_DK_RNV45_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV45
    Call ResultAdd("DK_RNV45", tmp183)

' #### DK_RNV46_S2 ####

    ' 0.複数画像情報インポート.DK_RNV46_S2

    ' 1.複数画像情報インポート.DK_RNV46_S2

    ' 2.Subtract(通常).DK_RNV46_S2

    ' 3.ExecuteLUT.DK_RNV46_S2

    ' 8.ShiftLeft.DK_RNV46_S2

    ' 9.複数画像用_LSB定義.DK_RNV46_S2

    ' 12.SliceLevel生成.DK_RNV46_S2
    Dim tmp_Slice47(nSite) As Double
    Call MakeSliceLevel(tmp_Slice47, 0.0047 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV46_S2
    Dim tmp184_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice47, tmp_Slice47, idpLimitEachSite, idpLimitExclude, tmp184_0)
    Dim tmp185 As CImgColorAllResult
    Call GetSum_CImgColor(tmp185, tmp184_0)

    ' 14.GetSum_Color.DK_RNV46_S2
    Dim tmp186(nSite) As Double
    Call GetSum_Color(tmp186, tmp185, "-")

    ' 15.PutTestResult.DK_RNV46_S2
    Call ResultAdd("DK_RNV46_S2", tmp186)

' #### DK_RNV46 ####

    ' 0.測定結果取得.DK_RNV46

    ' 1.測定結果取得.DK_RNV46
    Dim tmp_DK_RNV46_S2() As Double
    TheResult.GetResult "DK_RNV46_S2", tmp_DK_RNV46_S2

    ' 2.計算式評価.DK_RNV46
    Dim tmp187(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp187(site) = tmp_DK_RNV45_S2(site) - tmp_DK_RNV46_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV46
    Call ResultAdd("DK_RNV46", tmp187)

' #### DK_RNV47_S2 ####

    ' 0.複数画像情報インポート.DK_RNV47_S2

    ' 1.複数画像情報インポート.DK_RNV47_S2

    ' 2.Subtract(通常).DK_RNV47_S2

    ' 3.ExecuteLUT.DK_RNV47_S2

    ' 8.ShiftLeft.DK_RNV47_S2

    ' 9.複数画像用_LSB定義.DK_RNV47_S2

    ' 12.SliceLevel生成.DK_RNV47_S2
    Dim tmp_Slice48(nSite) As Double
    Call MakeSliceLevel(tmp_Slice48, 0.0048 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV47_S2
    Dim tmp188_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice48, tmp_Slice48, idpLimitEachSite, idpLimitExclude, tmp188_0)
    Dim tmp189 As CImgColorAllResult
    Call GetSum_CImgColor(tmp189, tmp188_0)

    ' 14.GetSum_Color.DK_RNV47_S2
    Dim tmp190(nSite) As Double
    Call GetSum_Color(tmp190, tmp189, "-")

    ' 15.PutTestResult.DK_RNV47_S2
    Call ResultAdd("DK_RNV47_S2", tmp190)

' #### DK_RNV47 ####

    ' 0.測定結果取得.DK_RNV47

    ' 1.測定結果取得.DK_RNV47
    Dim tmp_DK_RNV47_S2() As Double
    TheResult.GetResult "DK_RNV47_S2", tmp_DK_RNV47_S2

    ' 2.計算式評価.DK_RNV47
    Dim tmp191(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp191(site) = tmp_DK_RNV46_S2(site) - tmp_DK_RNV47_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV47
    Call ResultAdd("DK_RNV47", tmp191)

' #### DK_RNV48_S2 ####

    ' 0.複数画像情報インポート.DK_RNV48_S2

    ' 1.複数画像情報インポート.DK_RNV48_S2

    ' 2.Subtract(通常).DK_RNV48_S2

    ' 3.ExecuteLUT.DK_RNV48_S2

    ' 8.ShiftLeft.DK_RNV48_S2

    ' 9.複数画像用_LSB定義.DK_RNV48_S2

    ' 12.SliceLevel生成.DK_RNV48_S2
    Dim tmp_Slice49(nSite) As Double
    Call MakeSliceLevel(tmp_Slice49, 0.0049 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV48_S2
    Dim tmp192_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice49, tmp_Slice49, idpLimitEachSite, idpLimitExclude, tmp192_0)
    Dim tmp193 As CImgColorAllResult
    Call GetSum_CImgColor(tmp193, tmp192_0)

    ' 14.GetSum_Color.DK_RNV48_S2
    Dim tmp194(nSite) As Double
    Call GetSum_Color(tmp194, tmp193, "-")

    ' 15.PutTestResult.DK_RNV48_S2
    Call ResultAdd("DK_RNV48_S2", tmp194)

' #### DK_RNV48 ####

    ' 0.測定結果取得.DK_RNV48

    ' 1.測定結果取得.DK_RNV48
    Dim tmp_DK_RNV48_S2() As Double
    TheResult.GetResult "DK_RNV48_S2", tmp_DK_RNV48_S2

    ' 2.計算式評価.DK_RNV48
    Dim tmp195(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp195(site) = tmp_DK_RNV47_S2(site) - tmp_DK_RNV48_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV48
    Call ResultAdd("DK_RNV48", tmp195)

' #### DK_RNV50 ####

    ' 0.複数画像情報インポート.DK_RNV50

    ' 1.複数画像情報インポート.DK_RNV50

    ' 2.Subtract(通常).DK_RNV50

    ' 3.ExecuteLUT.DK_RNV50

    ' 8.ShiftLeft.DK_RNV50

    ' 9.複数画像用_LSB定義.DK_RNV50

    ' 12.SliceLevel生成.DK_RNV50
    Dim tmp_Slice50(nSite) As Double
    Call MakeSliceLevel(tmp_Slice50, 0.005 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNV50
    Dim tmp196_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice50, tmp_Slice50, idpLimitEachSite, idpLimitExclude, tmp196_0)
    Dim tmp197 As CImgColorAllResult
    Call GetSum_CImgColor(tmp197, tmp196_0)

    ' 14.GetSum_Color.DK_RNV50
    Dim tmp198(nSite) As Double
    Call GetSum_Color(tmp198, tmp197, "-")

    ' 15.PutTestResult.DK_RNV50
    Call ResultAdd("DK_RNV50", tmp198)

' #### DK_RNV49 ####

    ' 0.測定結果取得.DK_RNV49

    ' 1.測定結果取得.DK_RNV49
    Dim tmp_DK_RNV50() As Double
    TheResult.GetResult "DK_RNV50", tmp_DK_RNV50

    ' 2.計算式評価.DK_RNV49
    Dim tmp199(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp199(site) = tmp_DK_RNV48_S2(site) - tmp_DK_RNV50(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNV49
    Call ResultAdd("DK_RNV49", tmp199)

' #### DK_RNSGM ####

    ' 0.複数画像情報インポート.DK_RNSGM

    ' 1.複数画像情報インポート.DK_RNSGM

    ' 2.Subtract(通常).DK_RNSGM

    ' 3.ExecuteLUT.DK_RNSGM

    ' 4.Multiply(通常).DK_RNSGM
    Dim sPlane4 As CImgPlane
    Call GetFreePlane(sPlane4, "Normal_Bayer2x4", idpDepthS16, , "sPlane4")
    Call Multiply(sPlane2, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, sPlane2, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, sPlane4, "Bayer2x4_ZONE2D", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane2)

    ' 5.Average_FA.DK_RNSGM
    Dim tmp200_0 As CImgColorAllResult
    Call Average_FA(sPlane4, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, tmp200_0)
        Call ReleasePlane(sPlane4)
    Dim tmp201 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp201, tmp200_0)

    ' 6.GetAverage_Color.DK_RNSGM
    Dim tmp202(nSite) As Double
    Call GetAverage_Color(tmp202, tmp201, "-")

    ' 7.計算式評価.DK_RNSGM
    Dim tmp203(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp203(site) = Sqr(tmp202(site))
        End If
    Next site

    ' 9.複数画像用_LSB定義.DK_RNSGM

    ' 11.PutTestResult.DK_RNSGM
    Call ResultAdd("DK_RNSGM", tmp203)

' #### DK_RNL1_S1 ####

    ' 0.複数画像情報インポート.DK_RNL1_S1

    ' 1.複数画像情報インポート.DK_RNL1_S1

    ' 2.Subtract(通常).DK_RNL1_S1

    ' 3.ExecuteLUT.DK_RNL1_S1

    ' 8.ShiftLeft.DK_RNL1_S1

    ' 9.複数画像用_LSB定義.DK_RNL1_S1

    ' 12.SliceLevel生成.DK_RNL1_S1
    Dim tmp_Slice51(nSite) As Double
    Call MakeSliceLevel(tmp_Slice51, 0.0053 * Sqr(2) * 2 ^ 1, DK18_RNERR_LSB, , , , idpCountAbove)

    ' 13.Count_FA.DK_RNL1_S1
    Dim tmp204_0 As CImgColorAllResult
    Call count_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, idpCountAbove, tmp_Slice51, tmp_Slice51, idpLimitEachSite, idpLimitExclude, tmp204_0)
        Call ReleasePlane(sPlane3)
    Dim tmp205 As CImgColorAllResult
    Call GetSum_CImgColor(tmp205, tmp204_0)

    ' 14.GetSum_Color.DK_RNL1_S1
    Dim tmp206(nSite) As Double
    Call GetSum_Color(tmp206, tmp205, "-")

    ' 15.PutTestResult.DK_RNL1_S1
    Call ResultAdd("DK_RNL1_S1", tmp206)

' #### DK_RNL0 ####

    ' 0.測定結果取得.DK_RNL0

    ' 1.測定結果取得.DK_RNL0

    ' 2.計算式評価.DK_RNL0
    Dim tmp207(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp207(site) = tmp_DK_RNV13_S2(site) - tmp_DK_RNV27_S2(site)
        End If
    Next site

    ' 3.PutTestResult.DK_RNL0
    Call ResultAdd("DK_RNL0", tmp207)

' #### DK_RNL0_1M ####

    ' 0.GetResult.DK_RNL0_1M
    Dim tmp_DK_RNL0() As Double
     TheResult.GetResult "DK_RNL0", tmp_DK_RNL0

    ' 1.計算式評価.DK_RNL0_1M
    Dim tmp208(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp208(site) = TheIDP.PMD("Bayer2x4_ZONE2D").width
        End If
    Next site

    ' 2.計算式評価.DK_RNL0_1M
    Dim tmp209(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp209(site) = TheIDP.PMD("Bayer2x4_ZONE2D").height
        End If
    Next site

    ' 3.計算式評価.DK_RNL0_1M
    Dim tmp210(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp210(site) = 1000000
        End If
    Next site

    ' 4.計算式評価.DK_RNL0_1M
    Dim tmp211(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp211(site) = Div(tmp210(site), tmp208(site) * tmp209(site), 999)
        End If
    Next site

    ' 5.計算式評価.DK_RNL0_1M
    Dim tmp212(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp212(site) = tmp211(site) * tmp_DK_RNL0(site)
        End If
    Next site

    ' 6.PutTestResult.DK_RNL0_1M
    Call ResultAdd("DK_RNL0_1M", tmp212)

End Function



Attribute VB_Name = "Image_021_PMG18_ERR_Mod"

Option Explicit

Public Function PMG18_ERR_Process()

        Call PutImageInto_Common

' #### PMG18_CL ####

    Dim site As Long

    ' 0.画像情報インポート.PMG18_CL
    Dim PMG18_ERR_Param As CParamPlane
    Dim PMG18_ERR_DevInfo As CDeviceConfigInfo
    Dim PMG18_ERR_Plane As CImgPlane
    Set PMG18_ERR_Param = TheParameterBank.Item("PMG18ImageTest_Acq1")
    Set PMG18_ERR_DevInfo = TheDeviceProfiler.ConfigInfo("PMG18ImageTest_Acq1")
        Call TheParameterBank.Delete("PMG18ImageTest_Acq1")
    Set PMG18_ERR_Plane = PMG18_ERR_Param.plane

    ' 21.LSB定義.PMG18_CL
    Dim PMG18_ERR_LSB() As Double
     PMG18_ERR_LSB = PMG18_ERR_DevInfo.Lsb.AsDouble

    ' 22.SliceLevel生成.PMG18_CL
    Dim tmp_Slice1(nSite) As Double
    Call MakeSliceLevel(tmp_Slice1, 395)

    ' 23.SliceLevel生成.PMG18_CL
    Dim tmp_Slice2(nSite) As Double
    Call MakeSliceLevel(tmp_Slice2, 1)

    ' 24.Count_FA.PMG18_CL
    Dim tmp1_0 As CImgColorAllResult
    Call count_FA(PMG18_ERR_Plane, "Bayer2x4_ZONE3", EEE_COLOR_ALL, idpCountOutside, tmp_Slice2, tmp_Slice1, idpLimitEachSite, idpLimitExclude, tmp1_0, "Flg_Temp1")
    Dim tmp2 As CImgColorAllResult
    Call GetSum_CImgColor(tmp2, tmp1_0)

    ' 25.FlagCopy.PMG18_CL
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, True, "sPlane1")
    Call FlagCopy(sPlane1, "Bayer2x4_ZONE3", "Flg_Temp1", 1)
        Call ClearALLFlagBit("Flg_Temp1")

    ' 26.AccumulateColumn.PMG18_CL
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4_ACC", idpDepthS16, , "sPlane2")
    Call MakeAccPMD(sPlane2, "Bayer2x4_ZONE3", "Bayer2x4_ACC_ZONE3_EEE_COLOR_FLAT", EEE_COLOR_FLAT)
    Call AccumulateColumn(sPlane1, "Bayer2x4_ZONE3", EEE_COLOR_FLAT, sPlane2, "Bayer2x4_ACC_ZONE3_EEE_COLOR_FLAT", EEE_COLOR_FLAT, idpAccumSum)
        Call ReleasePlane(sPlane1)

    ' 27.Count_FA.PMG18_CL
    Dim tmp3_0 As CImgColorAllResult
    Call count_FA(sPlane2, "Bayer2x4_ACC_ZONE3_EEE_COLOR_FLAT", EEE_COLOR_ALL, idpCountAbove, 0, 0, idpLimitEachSite, idpLimitExclude, tmp3_0)
    Dim tmp4 As CImgColorAllResult
    Call GetSum_CImgColor(tmp4, tmp3_0)

    ' 28.GetSum_Color.PMG18_CL
    Dim tmp5(nSite) As Double
    Call GetSum_Color(tmp5, tmp4, "-")

    ' 29.PutTestResult.PMG18_CL
    Call ResultAdd("PMG18_CL", tmp5)

' #### PMG18_PXMX ####

    ' 0.画像情報インポート.PMG18_PXMX

    ' 21.LSB定義.PMG18_PXMX

    ' 22.SliceLevel生成.PMG18_PXMX

    ' 23.SliceLevel生成.PMG18_PXMX

    ' 24.Count_FA.PMG18_PXMX

    ' 25.FlagCopy.PMG18_PXMX

    ' 26.AccumulateColumn.PMG18_PXMX

    ' 30.Max_FA.PMG18_PXMX
    Dim tmp6_0 As CImgColorAllResult
    Call Max_FA(sPlane2, "Bayer2x4_ACC_ZONE3_EEE_COLOR_FLAT", EEE_COLOR_ALL, tmp6_0)
        Call ReleasePlane(sPlane2)
    Dim tmp7 As CImgColorAllResult
    Call GetMax_CImgColor(tmp7, tmp6_0)

    ' 31.GetMax_Color.PMG18_PXMX
    Dim tmp8(nSite) As Double
    Call GetMax_Color(tmp8, tmp7, "-")

    ' 32.PutTestResult.PMG18_PXMX
    Call ResultAdd("PMG18_PXMX", tmp8)

' #### PMG18_PX ####

    ' 33.GetSum_Color.PMG18_PX
    Dim tmp9(nSite) As Double
    Call GetSum_Color(tmp9, tmp2, "-")

    ' 34.PutTestResult.PMG18_PX
    Call ResultAdd("PMG18_PX", tmp9)

' #### PMG18_CLAV ####

    ' 0.画像情報インポート.PMG18_CLAV

    ' 33.Average_FA.PMG18_CLAV
    Dim tmp10_0 As CImgColorAllResult
    Call Average_FA(PMG18_ERR_Plane, "Bayer2x4_ZONE3", EEE_COLOR_ALL, tmp10_0)
    Dim tmp11 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp11, tmp10_0)

    ' 34.GetAverage_Color.PMG18_CLAV
    Dim tmp12(nSite) As Double
    Call GetAverage_Color(tmp12, tmp11, "-")

    ' 85.LSB定義.PMG18_CLAV

    ' 91.PutTestResult.PMG18_CLAV
    Call ResultAdd("PMG18_CLAV", tmp12)

' #### PMG18_CLMX ####

    ' 0.画像情報インポート.PMG18_CLMX

    ' 2.AccumulateColumn.PMG18_CLMX
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4_ACC", idpDepthF32, , "sPlane3")
    Call MakeAccPMD(sPlane3, "Bayer2x4_ZONE3", "Bayer2x4_ACC_ZONE3_EEE_COLOR_FLAT", EEE_COLOR_FLAT)
    Call AccumulateColumn(PMG18_ERR_Plane, "Bayer2x4_ZONE3", EEE_COLOR_FLAT, sPlane3, "Bayer2x4_ACC_ZONE3_EEE_COLOR_FLAT", EEE_COLOR_FLAT, idpAccumMean)

    ' 6.Max_FA.PMG18_CLMX
    Dim tmp13_0 As CImgColorAllResult
    Call Max_FA(sPlane3, "Bayer2x4_ACC_ZONE3_EEE_COLOR_FLAT", EEE_COLOR_ALL, tmp13_0)
    Dim tmp14 As CImgColorAllResult
    Call GetMax_CImgColor(tmp14, tmp13_0)

    ' 7.GetMax_Color.PMG18_CLMX
    Dim tmp15(nSite) As Double
    Call GetMax_Color(tmp15, tmp14, "-")

    ' 8.PutTestResult.PMG18_CLMX
    Call ResultAdd("PMG18_CLMX", tmp15)

' #### PMG18_CLMN ####

    ' 0.画像情報インポート.PMG18_CLMN

    ' 2.AccumulateColumn.PMG18_CLMN

    ' 3.Min_FA.PMG18_CLMN
    Dim tmp16_0 As CImgColorAllResult
    Call Min_FA(sPlane3, "Bayer2x4_ACC_ZONE3_EEE_COLOR_FLAT", EEE_COLOR_ALL, tmp16_0)
        Call ReleasePlane(sPlane3)
    Dim tmp17 As CImgColorAllResult
    Call GetMin_CImgColor(tmp17, tmp16_0)

    ' 4.GetMin_Color.PMG18_CLMN
    Dim tmp18(nSite) As Double
    Call GetMin_Color(tmp18, tmp17, "-")

    ' 5.PutTestResult.PMG18_CLMN
    Call ResultAdd("PMG18_CLMN", tmp18)

End Function



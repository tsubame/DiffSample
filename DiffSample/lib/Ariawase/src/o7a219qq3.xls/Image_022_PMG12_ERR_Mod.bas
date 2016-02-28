Attribute VB_Name = "Image_022_PMG12_ERR_Mod"

Option Explicit

Public Function PMG12_ERR_Process()

        Call PutImageInto_Common

' #### PMG12_CL ####

    Dim site As Long

    ' 0.画像情報インポート.PMG12_CL
    Dim PMG12_ERR_Param As CParamPlane
    Dim PMG12_ERR_DevInfo As CDeviceConfigInfo
    Dim PMG12_ERR_Plane As CImgPlane
    Set PMG12_ERR_Param = TheParameterBank.Item("PMG12ImageTest_Acq1")
    Set PMG12_ERR_DevInfo = TheDeviceProfiler.ConfigInfo("PMG12ImageTest_Acq1")
        Call TheParameterBank.Delete("PMG12ImageTest_Acq1")
    Set PMG12_ERR_Plane = PMG12_ERR_Param.plane

    ' 21.LSB定義.PMG12_CL
    Dim PMG12_ERR_LSB() As Double
     PMG12_ERR_LSB = PMG12_ERR_DevInfo.Lsb.AsDouble

    ' 22.SliceLevel生成.PMG12_CL
    Dim tmp_Slice1(nSite) As Double
    Call MakeSliceLevel(tmp_Slice1, 395)

    ' 23.SliceLevel生成.PMG12_CL
    Dim tmp_Slice2(nSite) As Double
    Call MakeSliceLevel(tmp_Slice2, 1)

    ' 24.Count_FA.PMG12_CL
    Dim tmp1_0 As CImgColorAllResult
    Call count_FA(PMG12_ERR_Plane, "Bayer2x4_ZONE3", EEE_COLOR_ALL, idpCountOutside, tmp_Slice2, tmp_Slice1, idpLimitEachSite, idpLimitExclude, tmp1_0, "Flg_Temp1")
    Dim tmp2 As CImgColorAllResult
    Call GetSum_CImgColor(tmp2, tmp1_0)

    ' 25.FlagCopy.PMG12_CL
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, True, "sPlane1")
    Call FlagCopy(sPlane1, "Bayer2x4_ZONE3", "Flg_Temp1", 1)
        Call ClearALLFlagBit("Flg_Temp1")

    ' 26.AccumulateColumn.PMG12_CL
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4_ACC", idpDepthS16, , "sPlane2")
    Call MakeAccPMD(sPlane2, "Bayer2x4_ZONE3", "Bayer2x4_ACC_ZONE3_EEE_COLOR_FLAT", EEE_COLOR_FLAT)
    Call AccumulateColumn(sPlane1, "Bayer2x4_ZONE3", EEE_COLOR_FLAT, sPlane2, "Bayer2x4_ACC_ZONE3_EEE_COLOR_FLAT", EEE_COLOR_FLAT, idpAccumSum)
        Call ReleasePlane(sPlane1)

    ' 27.Count_FA.PMG12_CL
    Dim tmp3_0 As CImgColorAllResult
    Call count_FA(sPlane2, "Bayer2x4_ACC_ZONE3_EEE_COLOR_FLAT", EEE_COLOR_ALL, idpCountAbove, 0, 0, idpLimitEachSite, idpLimitExclude, tmp3_0)
    Dim tmp4 As CImgColorAllResult
    Call GetSum_CImgColor(tmp4, tmp3_0)

    ' 28.GetSum_Color.PMG12_CL
    Dim tmp5(nSite) As Double
    Call GetSum_Color(tmp5, tmp4, "-")

    ' 29.PutTestResult.PMG12_CL
    Call ResultAdd("PMG12_CL", tmp5)

' #### PMG12_PXMX ####

    ' 0.画像情報インポート.PMG12_PXMX

    ' 21.LSB定義.PMG12_PXMX

    ' 22.SliceLevel生成.PMG12_PXMX

    ' 23.SliceLevel生成.PMG12_PXMX

    ' 24.Count_FA.PMG12_PXMX

    ' 25.FlagCopy.PMG12_PXMX

    ' 26.AccumulateColumn.PMG12_PXMX

    ' 30.Max_FA.PMG12_PXMX
    Dim tmp6_0 As CImgColorAllResult
    Call Max_FA(sPlane2, "Bayer2x4_ACC_ZONE3_EEE_COLOR_FLAT", EEE_COLOR_ALL, tmp6_0)
        Call ReleasePlane(sPlane2)
    Dim tmp7 As CImgColorAllResult
    Call GetMax_CImgColor(tmp7, tmp6_0)

    ' 31.GetMax_Color.PMG12_PXMX
    Dim tmp8(nSite) As Double
    Call GetMax_Color(tmp8, tmp7, "-")

    ' 32.PutTestResult.PMG12_PXMX
    Call ResultAdd("PMG12_PXMX", tmp8)

' #### PMG12_PX ####

    ' 33.GetSum_Color.PMG12_PX
    Dim tmp9(nSite) As Double
    Call GetSum_Color(tmp9, tmp2, "-")

    ' 34.PutTestResult.PMG12_PX
    Call ResultAdd("PMG12_PX", tmp9)

' #### PMG12_CLAV ####

    ' 0.画像情報インポート.PMG12_CLAV

    ' 33.Average_FA.PMG12_CLAV
    Dim tmp10_0 As CImgColorAllResult
    Call Average_FA(PMG12_ERR_Plane, "Bayer2x4_ZONE3", EEE_COLOR_ALL, tmp10_0)
    Dim tmp11 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp11, tmp10_0)

    ' 34.GetAverage_Color.PMG12_CLAV
    Dim tmp12(nSite) As Double
    Call GetAverage_Color(tmp12, tmp11, "-")

    ' 85.LSB定義.PMG12_CLAV

    ' 91.PutTestResult.PMG12_CLAV
    Call ResultAdd("PMG12_CLAV", tmp12)

' #### PMG12_CLMX ####

    ' 0.画像情報インポート.PMG12_CLMX

    ' 2.AccumulateColumn.PMG12_CLMX
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4_ACC", idpDepthF32, , "sPlane3")
    Call MakeAccPMD(sPlane3, "Bayer2x4_ZONE3", "Bayer2x4_ACC_ZONE3_EEE_COLOR_FLAT", EEE_COLOR_FLAT)
    Call AccumulateColumn(PMG12_ERR_Plane, "Bayer2x4_ZONE3", EEE_COLOR_FLAT, sPlane3, "Bayer2x4_ACC_ZONE3_EEE_COLOR_FLAT", EEE_COLOR_FLAT, idpAccumMean)

    ' 6.Max_FA.PMG12_CLMX
    Dim tmp13_0 As CImgColorAllResult
    Call Max_FA(sPlane3, "Bayer2x4_ACC_ZONE3_EEE_COLOR_FLAT", EEE_COLOR_ALL, tmp13_0)
    Dim tmp14 As CImgColorAllResult
    Call GetMax_CImgColor(tmp14, tmp13_0)

    ' 7.GetMax_Color.PMG12_CLMX
    Dim tmp15(nSite) As Double
    Call GetMax_Color(tmp15, tmp14, "-")

    ' 8.PutTestResult.PMG12_CLMX
    Call ResultAdd("PMG12_CLMX", tmp15)

' #### PMG12_CLMN ####

    ' 0.画像情報インポート.PMG12_CLMN

    ' 2.AccumulateColumn.PMG12_CLMN

    ' 3.Min_FA.PMG12_CLMN
    Dim tmp16_0 As CImgColorAllResult
    Call Min_FA(sPlane3, "Bayer2x4_ACC_ZONE3_EEE_COLOR_FLAT", EEE_COLOR_ALL, tmp16_0)
        Call ReleasePlane(sPlane3)
    Dim tmp17 As CImgColorAllResult
    Call GetMin_CImgColor(tmp17, tmp16_0)

    ' 4.GetMin_Color.PMG12_CLMN
    Dim tmp18(nSite) As Double
    Call GetMin_Color(tmp18, tmp17, "-")

    ' 5.PutTestResult.PMG12_CLMN
    Call ResultAdd("PMG12_CLMN", tmp18)

End Function



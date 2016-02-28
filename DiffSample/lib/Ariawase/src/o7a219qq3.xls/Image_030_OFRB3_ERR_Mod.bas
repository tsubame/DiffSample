Attribute VB_Name = "Image_030_OFRB3_ERR_Mod"

Option Explicit

Public Function OFRB3_ERR_Process()

        Call PutImageInto_Common

' #### OFRB3_SENR ####

    Dim site As Long

    ' 0.画像情報インポート.OFRB3_SENR
    Dim OFRB3_ERR_Param As CParamPlane
    Dim OFRB3_ERR_DevInfo As CDeviceConfigInfo
    Dim OFRB3_ERR_Plane As CImgPlane
    Set OFRB3_ERR_Param = TheParameterBank.Item("OFRB3ImageTest_Acq1")
    Set OFRB3_ERR_DevInfo = TheDeviceProfiler.ConfigInfo("OFRB3ImageTest_Acq1")
        Call TheParameterBank.Delete("OFRB3ImageTest_Acq1")
    Set OFRB3_ERR_Plane = OFRB3_ERR_Param.plane

    ' 1.Clamp.OFRB3_SENR
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call Clamp(OFRB3_ERR_Plane, sPlane1, "Bayer2x4_VOPB")

    ' 2.Median.OFRB3_SENR
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 1, 5)
        Call ReleasePlane(sPlane1)

    ' 3.Median.OFRB3_SENR
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call MedianEx(sPlane2, sPlane3, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane2)

    ' 82.Average_FA.OFRB3_SENR
    Dim tmp1_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONE0", EEE_COLOR_ALL, tmp1_0)
        Call ReleasePlane(sPlane3)
    Dim tmp2 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp2, tmp1_0)

    ' 83.GetAverage_Color.OFRB3_SENR
    Dim tmp3(nSite) As Double
    Call GetAverage_Color(tmp3, tmp2, "R1", "R2")

    ' 233.LSB定義.OFRB3_SENR
    Dim OFRB3_ERR_LSB() As Double
     OFRB3_ERR_LSB = OFRB3_ERR_DevInfo.Lsb.AsDouble

    ' 238.LSB換算.OFRB3_SENR
    Dim tmp4(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp4(site) = tmp3(site) * OFRB3_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.OFRB3_SENR
    Call ResultAdd("OFRB3_SENR", tmp4)

' #### OFRB3_SENGR ####

    ' 0.画像情報インポート.OFRB3_SENGR

    ' 1.Clamp.OFRB3_SENGR

    ' 2.Median.OFRB3_SENGR

    ' 3.Median.OFRB3_SENGR

    ' 82.Average_FA.OFRB3_SENGR

    ' 83.GetAverage_Color.OFRB3_SENGR
    Dim tmp5(nSite) As Double
    Call GetAverage_Color(tmp5, tmp2, "Gr1", "Gr2")

    ' 233.LSB定義.OFRB3_SENGR

    ' 238.LSB換算.OFRB3_SENGR
    Dim tmp6(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp6(site) = tmp5(site) * OFRB3_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.OFRB3_SENGR
    Call ResultAdd("OFRB3_SENGR", tmp6)

' #### OFRB3_SENGB ####

    ' 0.画像情報インポート.OFRB3_SENGB

    ' 1.Clamp.OFRB3_SENGB

    ' 2.Median.OFRB3_SENGB

    ' 3.Median.OFRB3_SENGB

    ' 82.Average_FA.OFRB3_SENGB

    ' 83.GetAverage_Color.OFRB3_SENGB
    Dim tmp7(nSite) As Double
    Call GetAverage_Color(tmp7, tmp2, "Gb1", "Gb2")

    ' 233.LSB定義.OFRB3_SENGB

    ' 238.LSB換算.OFRB3_SENGB
    Dim tmp8(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp8(site) = tmp7(site) * OFRB3_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.OFRB3_SENGB
    Call ResultAdd("OFRB3_SENGB", tmp8)

' #### OFR_BLMGR ####

    ' 0.画像情報インポート.OFR_BLMGR

    ' 1.Clamp.OFR_BLMGR

    ' 2.Median.OFR_BLMGR

    ' 3.Median.OFR_BLMGR

    ' 5.Average_FA.OFR_BLMGR

    ' 6.GetAverage_Color.OFR_BLMGR

    ' 7.画像情報インポート.OFR_BLMGR
    Dim OFRB2_ERR_Param As CParamPlane
    Dim OFRB2_ERR_DevInfo As CDeviceConfigInfo
    Dim OFRB2_ERR_Plane As CImgPlane
    Set OFRB2_ERR_Param = TheParameterBank.Item("OFRB2ImageTest_Acq1")
    Set OFRB2_ERR_DevInfo = TheDeviceProfiler.ConfigInfo("OFRB2ImageTest_Acq1")
        Call TheParameterBank.Delete("OFRB2ImageTest_Acq1")
    Set OFRB2_ERR_Plane = OFRB2_ERR_Param.plane

    ' 8.Clamp.OFR_BLMGR
    Dim sPlane4 As CImgPlane
    Call GetFreePlane(sPlane4, "Normal_Bayer2x4", idpDepthS16, , "sPlane4")
    Call Clamp(OFRB2_ERR_Plane, sPlane4, "Bayer2x4_VOPB")

    ' 9.Median.OFR_BLMGR
    Dim sPlane5 As CImgPlane
    Call GetFreePlane(sPlane5, "Normal_Bayer2x4", idpDepthS16, , "sPlane5")
    Call MedianEx(sPlane4, sPlane5, "Bayer2x4_ZONE3", 1, 5)
        Call ReleasePlane(sPlane4)

    ' 10.Median.OFR_BLMGR
    Dim sPlane6 As CImgPlane
    Call GetFreePlane(sPlane6, "Normal_Bayer2x4", idpDepthS16, , "sPlane6")
    Call MedianEx(sPlane5, sPlane6, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane5)

    ' 12.Average_FA.OFR_BLMGR
    Dim tmp9_0 As CImgColorAllResult
    Call Average_FA(sPlane6, "Bayer2x4_ZONE0", EEE_COLOR_ALL, tmp9_0)
        Call ReleasePlane(sPlane6)
    Dim tmp10 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp10, tmp9_0)

    ' 13.GetAverage_Color.OFR_BLMGR
    Dim tmp11(nSite) As Double
    Call GetAverage_Color(tmp11, tmp10, "Gr1", "Gr2")

    ' 14.画像情報インポート.OFR_BLMGR
    Dim OFRB1_ERR_Param As CParamPlane
    Dim OFRB1_ERR_DevInfo As CDeviceConfigInfo
    Dim OFRB1_ERR_Plane As CImgPlane
    Set OFRB1_ERR_Param = TheParameterBank.Item("OFRB1ImageTest_Acq1")
    Set OFRB1_ERR_DevInfo = TheDeviceProfiler.ConfigInfo("OFRB1ImageTest_Acq1")
        Call TheParameterBank.Delete("OFRB1ImageTest_Acq1")
    Set OFRB1_ERR_Plane = OFRB1_ERR_Param.plane

    ' 15.Clamp.OFR_BLMGR
    Dim sPlane7 As CImgPlane
    Call GetFreePlane(sPlane7, "Normal_Bayer2x4", idpDepthS16, , "sPlane7")
    Call Clamp(OFRB1_ERR_Plane, sPlane7, "Bayer2x4_VOPB")

    ' 16.Median.OFR_BLMGR
    Dim sPlane8 As CImgPlane
    Call GetFreePlane(sPlane8, "Normal_Bayer2x4", idpDepthS16, , "sPlane8")
    Call MedianEx(sPlane7, sPlane8, "Bayer2x4_ZONE3", 1, 5)
        Call ReleasePlane(sPlane7)

    ' 17.Median.OFR_BLMGR
    Dim sPlane9 As CImgPlane
    Call GetFreePlane(sPlane9, "Normal_Bayer2x4", idpDepthS16, , "sPlane9")
    Call MedianEx(sPlane8, sPlane9, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane8)

    ' 19.Average_FA.OFR_BLMGR
    Dim tmp12_0 As CImgColorAllResult
    Call Average_FA(sPlane9, "Bayer2x4_ZONE0", EEE_COLOR_ALL, tmp12_0)
        Call ReleasePlane(sPlane9)
    Dim tmp13 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp13, tmp12_0)

    ' 20.GetAverage_Color.OFR_BLMGR
    Dim tmp14(nSite) As Double
    Call GetAverage_Color(tmp14, tmp13, "R1", "R2")

    ' 21.GetAverage_Color.OFR_BLMGR
    Dim tmp15(nSite) As Double
    Call GetAverage_Color(tmp15, tmp13, "Gr1", "Gr2")

    ' 22.LSB定義.OFR_BLMGR

    ' 24.LSB定義.OFR_BLMGR
    Dim OFRB2_ERR_LSB() As Double
     OFRB2_ERR_LSB = OFRB2_ERR_DevInfo.Lsb.AsDouble

    ' 26.LSB定義.OFR_BLMGR
    Dim OFRB1_ERR_LSB() As Double
     OFRB1_ERR_LSB = OFRB1_ERR_DevInfo.Lsb.AsDouble

    ' 29.AccTime定義.OFR_BLMGR
    Dim OFRB3_ERR_AccTime() As Double
    OFRB3_ERR_AccTime = OFRB3_ERR_DevInfo.AccTime.AsAccTimeH

    ' 30.AccTime定義.OFR_BLMGR
    Dim OFRB2_ERR_AccTime() As Double
    OFRB2_ERR_AccTime = OFRB2_ERR_DevInfo.AccTime.AsAccTimeH

    ' 31.AccTime定義.OFR_BLMGR
    Dim OFRB1_ERR_AccTime() As Double
    OFRB1_ERR_AccTime = OFRB1_ERR_DevInfo.AccTime.AsAccTimeH

    ' 32.計算式評価.OFR_BLMGR
    Dim tmp16(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp16(site) = Div(OFRB3_ERR_AccTime(site) - OFRB2_ERR_AccTime(site), OFRB1_ERR_AccTime(site), 999)
        End If
    Next site

    ' 33.計算式評価.OFR_BLMGR
    Dim tmp17(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp17(site) = Div((tmp5(site) - tmp11(site) - (tmp15(site) * tmp16(site))), tmp14(site) * tmp16(site), 999)
        End If
    Next site

    ' 34.PutTestResult.OFR_BLMGR
    Call ResultAdd("OFR_BLMGR", tmp17)

' #### OFR_BLMGB ####

    ' 0.画像情報インポート.OFR_BLMGB

    ' 1.Clamp.OFR_BLMGB

    ' 2.Median.OFR_BLMGB

    ' 3.Median.OFR_BLMGB

    ' 5.Average_FA.OFR_BLMGB

    ' 6.GetAverage_Color.OFR_BLMGB

    ' 7.画像情報インポート.OFR_BLMGB

    ' 8.Clamp.OFR_BLMGB

    ' 9.Median.OFR_BLMGB

    ' 10.Median.OFR_BLMGB

    ' 12.Average_FA.OFR_BLMGB

    ' 13.GetAverage_Color.OFR_BLMGB
    Dim tmp18(nSite) As Double
    Call GetAverage_Color(tmp18, tmp10, "Gb1", "Gb2")

    ' 14.画像情報インポート.OFR_BLMGB

    ' 15.Clamp.OFR_BLMGB

    ' 16.Median.OFR_BLMGB

    ' 17.Median.OFR_BLMGB

    ' 19.Average_FA.OFR_BLMGB

    ' 20.GetAverage_Color.OFR_BLMGB

    ' 21.GetAverage_Color.OFR_BLMGB
    Dim tmp19(nSite) As Double
    Call GetAverage_Color(tmp19, tmp13, "Gb1", "Gb2")

    ' 22.LSB定義.OFR_BLMGB

    ' 24.LSB定義.OFR_BLMGB

    ' 26.LSB定義.OFR_BLMGB

    ' 29.AccTime定義.OFR_BLMGB

    ' 30.AccTime定義.OFR_BLMGB

    ' 31.AccTime定義.OFR_BLMGB

    ' 32.計算式評価.OFR_BLMGB

    ' 33.計算式評価.OFR_BLMGB
    Dim tmp20(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp20(site) = Div((tmp7(site) - tmp18(site) - (tmp19(site) * tmp16(site))), tmp14(site) * tmp16(site), 999)
        End If
    Next site

    ' 34.PutTestResult.OFR_BLMGB
    Call ResultAdd("OFR_BLMGB", tmp20)

End Function



Attribute VB_Name = "Image_023_OF_MGERR_Mod"

Option Explicit

Public Function OF_MGERR_Process()

        Call PutImageInto_Common

' #### OF_QSMNR1 ####

    Dim site As Long

    ' 0.画像情報インポート.OF_QSMNR1
    Dim OF_MGERR_Param As CParamPlane
    Dim OF_MGERR_DevInfo As CDeviceConfigInfo
    Dim OF_MGERR_Plane As CImgPlane
    Set OF_MGERR_Param = TheParameterBank.Item("OF_MGImageTest_Acq1")
    Set OF_MGERR_DevInfo = TheDeviceProfiler.ConfigInfo("OF_MGImageTest_Acq1")
        Call TheParameterBank.Delete("OF_MGImageTest_Acq1")
    Set OF_MGERR_Plane = OF_MGERR_Param.plane

    ' 1.Clamp.OF_QSMNR1
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call Clamp(OF_MGERR_Plane, sPlane1, "Bayer2x4_VOPB")

    ' 2.Median.OF_QSMNR1
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 1, 5)
        Call ReleasePlane(sPlane1)

    ' 3.Median.OF_QSMNR1
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call MedianEx(sPlane2, sPlane3, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane2)

    ' 45.Min_FA.OF_QSMNR1
    Dim tmp1_0 As CImgColorAllResult
    Call Min_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, tmp1_0)
    Dim tmp2 As CImgColorAllResult
    Call GetMin_CImgColor(tmp2, tmp1_0)

    ' 46.GetMin_Color.OF_QSMNR1
    Dim tmp3(nSite) As Double
    Call GetMin_Color(tmp3, tmp2, "R1")

    ' 233.LSB定義.OF_QSMNR1
    Dim OF_MGERR_LSB() As Double
     OF_MGERR_LSB = OF_MGERR_DevInfo.Lsb.AsDouble

    ' 240.LSB換算.OF_QSMNR1
    Dim tmp4(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp4(site) = tmp3(site) * OF_MGERR_LSB(site)
        End If
    Next site

    ' 241.PutTestResult.OF_QSMNR1
    Call ResultAdd("OF_QSMNR1", tmp4)

' #### OF_QSMNGR1 ####

    ' 0.画像情報インポート.OF_QSMNGR1

    ' 1.Clamp.OF_QSMNGR1

    ' 2.Median.OF_QSMNGR1

    ' 3.Median.OF_QSMNGR1

    ' 45.Min_FA.OF_QSMNGR1

    ' 46.GetMin_Color.OF_QSMNGR1
    Dim tmp5(nSite) As Double
    Call GetMin_Color(tmp5, tmp2, "Gr1")

    ' 233.LSB定義.OF_QSMNGR1

    ' 240.LSB換算.OF_QSMNGR1
    Dim tmp6(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp6(site) = tmp5(site) * OF_MGERR_LSB(site)
        End If
    Next site

    ' 241.PutTestResult.OF_QSMNGR1
    Call ResultAdd("OF_QSMNGR1", tmp6)

' #### OF_QSMNGB1 ####

    ' 0.画像情報インポート.OF_QSMNGB1

    ' 1.Clamp.OF_QSMNGB1

    ' 2.Median.OF_QSMNGB1

    ' 3.Median.OF_QSMNGB1

    ' 45.Min_FA.OF_QSMNGB1

    ' 46.GetMin_Color.OF_QSMNGB1
    Dim tmp7(nSite) As Double
    Call GetMin_Color(tmp7, tmp2, "Gb1")

    ' 233.LSB定義.OF_QSMNGB1

    ' 240.LSB換算.OF_QSMNGB1
    Dim tmp8(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp8(site) = tmp7(site) * OF_MGERR_LSB(site)
        End If
    Next site

    ' 241.PutTestResult.OF_QSMNGB1
    Call ResultAdd("OF_QSMNGB1", tmp8)

' #### OF_QSMNB1 ####

    ' 0.画像情報インポート.OF_QSMNB1

    ' 1.Clamp.OF_QSMNB1

    ' 2.Median.OF_QSMNB1

    ' 3.Median.OF_QSMNB1

    ' 45.Min_FA.OF_QSMNB1

    ' 46.GetMin_Color.OF_QSMNB1
    Dim tmp9(nSite) As Double
    Call GetMin_Color(tmp9, tmp2, "B1")

    ' 233.LSB定義.OF_QSMNB1

    ' 240.LSB換算.OF_QSMNB1
    Dim tmp10(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp10(site) = tmp9(site) * OF_MGERR_LSB(site)
        End If
    Next site

    ' 241.PutTestResult.OF_QSMNB1
    Call ResultAdd("OF_QSMNB1", tmp10)

' #### OF_QSMNR2 ####

    ' 0.画像情報インポート.OF_QSMNR2

    ' 1.Clamp.OF_QSMNR2

    ' 2.Median.OF_QSMNR2

    ' 3.Median.OF_QSMNR2

    ' 45.Min_FA.OF_QSMNR2

    ' 46.GetMin_Color.OF_QSMNR2
    Dim tmp11(nSite) As Double
    Call GetMin_Color(tmp11, tmp2, "R2")

    ' 233.LSB定義.OF_QSMNR2

    ' 240.LSB換算.OF_QSMNR2
    Dim tmp12(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp12(site) = tmp11(site) * OF_MGERR_LSB(site)
        End If
    Next site

    ' 241.PutTestResult.OF_QSMNR2
    Call ResultAdd("OF_QSMNR2", tmp12)

' #### OF_QSMNGR2 ####

    ' 0.画像情報インポート.OF_QSMNGR2

    ' 1.Clamp.OF_QSMNGR2

    ' 2.Median.OF_QSMNGR2

    ' 3.Median.OF_QSMNGR2

    ' 45.Min_FA.OF_QSMNGR2

    ' 46.GetMin_Color.OF_QSMNGR2
    Dim tmp13(nSite) As Double
    Call GetMin_Color(tmp13, tmp2, "Gr2")

    ' 233.LSB定義.OF_QSMNGR2

    ' 240.LSB換算.OF_QSMNGR2
    Dim tmp14(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp14(site) = tmp13(site) * OF_MGERR_LSB(site)
        End If
    Next site

    ' 241.PutTestResult.OF_QSMNGR2
    Call ResultAdd("OF_QSMNGR2", tmp14)

' #### OF_QSMNGB2 ####

    ' 0.画像情報インポート.OF_QSMNGB2

    ' 1.Clamp.OF_QSMNGB2

    ' 2.Median.OF_QSMNGB2

    ' 3.Median.OF_QSMNGB2

    ' 45.Min_FA.OF_QSMNGB2

    ' 46.GetMin_Color.OF_QSMNGB2
    Dim tmp15(nSite) As Double
    Call GetMin_Color(tmp15, tmp2, "Gb2")

    ' 233.LSB定義.OF_QSMNGB2

    ' 240.LSB換算.OF_QSMNGB2
    Dim tmp16(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp16(site) = tmp15(site) * OF_MGERR_LSB(site)
        End If
    Next site

    ' 241.PutTestResult.OF_QSMNGB2
    Call ResultAdd("OF_QSMNGB2", tmp16)

' #### OF_QSMNB2 ####

    ' 0.画像情報インポート.OF_QSMNB2

    ' 1.Clamp.OF_QSMNB2

    ' 2.Median.OF_QSMNB2

    ' 3.Median.OF_QSMNB2

    ' 45.Min_FA.OF_QSMNB2

    ' 46.GetMin_Color.OF_QSMNB2
    Dim tmp17(nSite) As Double
    Call GetMin_Color(tmp17, tmp2, "B2")

    ' 233.LSB定義.OF_QSMNB2

    ' 240.LSB換算.OF_QSMNB2
    Dim tmp18(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp18(site) = tmp17(site) * OF_MGERR_LSB(site)
        End If
    Next site

    ' 241.PutTestResult.OF_QSMNB2
    Call ResultAdd("OF_QSMNB2", tmp18)

' #### OF_QSAV ####

    ' 0.画像情報インポート.OF_QSAV

    ' 1.Clamp.OF_QSAV

    ' 2.Median.OF_QSAV

    ' 3.Median.OF_QSAV

    ' 82.Average_FA.OF_QSAV
    Dim tmp19_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, tmp19_0)
    Dim tmp20 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp20, tmp19_0)

    ' 83.GetAverage_Color.OF_QSAV
    Dim tmp21(nSite) As Double
    Call GetAverage_Color(tmp21, tmp20, "-")

    ' 233.LSB定義.OF_QSAV

    ' 238.LSB換算.OF_QSAV
    Dim tmp22(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp22(site) = tmp21(site) * OF_MGERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.OF_QSAV
    Call ResultAdd("OF_QSAV", tmp22)

' #### OF_QSAVR ####

    ' 0.画像情報インポート.OF_QSAVR

    ' 1.Clamp.OF_QSAVR

    ' 2.Median.OF_QSAVR

    ' 3.Median.OF_QSAVR

    ' 82.Average_FA.OF_QSAVR

    ' 83.GetAverage_Color.OF_QSAVR
    Dim tmp23(nSite) As Double
    Call GetAverage_Color(tmp23, tmp20, "R1", "R2")

    ' 233.LSB定義.OF_QSAVR

    ' 238.LSB換算.OF_QSAVR
    Dim tmp24(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp24(site) = tmp23(site) * OF_MGERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.OF_QSAVR
    Call ResultAdd("OF_QSAVR", tmp24)

' #### OF_QSAVGR ####

    ' 0.画像情報インポート.OF_QSAVGR

    ' 1.Clamp.OF_QSAVGR

    ' 2.Median.OF_QSAVGR

    ' 3.Median.OF_QSAVGR

    ' 82.Average_FA.OF_QSAVGR

    ' 83.GetAverage_Color.OF_QSAVGR
    Dim tmp25(nSite) As Double
    Call GetAverage_Color(tmp25, tmp20, "Gr1", "Gr2")

    ' 233.LSB定義.OF_QSAVGR

    ' 238.LSB換算.OF_QSAVGR
    Dim tmp26(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp26(site) = tmp25(site) * OF_MGERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.OF_QSAVGR
    Call ResultAdd("OF_QSAVGR", tmp26)

' #### OF_QSAVGB ####

    ' 0.画像情報インポート.OF_QSAVGB

    ' 1.Clamp.OF_QSAVGB

    ' 2.Median.OF_QSAVGB

    ' 3.Median.OF_QSAVGB

    ' 82.Average_FA.OF_QSAVGB

    ' 83.GetAverage_Color.OF_QSAVGB
    Dim tmp27(nSite) As Double
    Call GetAverage_Color(tmp27, tmp20, "Gb1", "Gb2")

    ' 233.LSB定義.OF_QSAVGB

    ' 238.LSB換算.OF_QSAVGB
    Dim tmp28(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp28(site) = tmp27(site) * OF_MGERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.OF_QSAVGB
    Call ResultAdd("OF_QSAVGB", tmp28)

' #### OF_QSAVB ####

    ' 0.画像情報インポート.OF_QSAVB

    ' 1.Clamp.OF_QSAVB

    ' 2.Median.OF_QSAVB

    ' 3.Median.OF_QSAVB

    ' 82.Average_FA.OF_QSAVB

    ' 83.GetAverage_Color.OF_QSAVB
    Dim tmp29(nSite) As Double
    Call GetAverage_Color(tmp29, tmp20, "B1", "B2")

    ' 233.LSB定義.OF_QSAVB

    ' 238.LSB換算.OF_QSAVB
    Dim tmp30(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp30(site) = tmp29(site) * OF_MGERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.OF_QSAVB
    Call ResultAdd("OF_QSAVB", tmp30)

' #### OF_QSAV_Z0 ####

    ' 0.画像情報インポート.OF_QSAV_Z0

    ' 1.Clamp.OF_QSAV_Z0

    ' 2.Median.OF_QSAV_Z0

    ' 3.Median.OF_QSAV_Z0

    ' 82.Average_FA.OF_QSAV_Z0
    Dim tmp31_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONE0", EEE_COLOR_ALL, tmp31_0)
        Call ReleasePlane(sPlane3)
    Dim tmp32 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp32, tmp31_0)

    ' 83.GetAverage_Color.OF_QSAV_Z0
    Dim tmp33(nSite) As Double
    Call GetAverage_Color(tmp33, tmp32, "-")

    ' 233.LSB定義.OF_QSAV_Z0

    ' 238.LSB換算.OF_QSAV_Z0
    Dim tmp34(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp34(site) = tmp33(site) * OF_MGERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.OF_QSAV_Z0
    Call ResultAdd("OF_QSAV_Z0", tmp34)

End Function



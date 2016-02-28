Attribute VB_Name = "Image_026_HL_MGERR_Mod"

Option Explicit

Public Function HL_MGERR_Process()

        Call PutImageInto_Common

' #### HL_SENR1 ####

    Dim site As Long

    ' 0.画像情報インポート.HL_SENR1
    Dim HL_MGERR_Param As CParamPlane
    Dim HL_MGERR_DevInfo As CDeviceConfigInfo
    Dim HL_MGERR_Plane As CImgPlane
    Set HL_MGERR_Param = TheParameterBank.Item("HL_MGImageTest_Acq1")
    Set HL_MGERR_DevInfo = TheDeviceProfiler.ConfigInfo("HL_MGImageTest_Acq1")
    Set HL_MGERR_Plane = HL_MGERR_Param.plane

    ' 1.Clamp.HL_SENR1
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call Clamp(HL_MGERR_Plane, sPlane1, "Bayer2x4_VOPB")

    ' 2.Median.HL_SENR1
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 1, 5)
        Call ReleasePlane(sPlane1)

    ' 3.Median.HL_SENR1
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call MedianEx(sPlane2, sPlane3, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane2)

    ' 82.Average_FA.HL_SENR1
    Dim tmp1_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONE0", EEE_COLOR_ALL, tmp1_0)
        Call ReleasePlane(sPlane3)
    Dim tmp2 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp2, tmp1_0)

    ' 83.GetAverage_Color.HL_SENR1
    Dim tmp3(nSite) As Double
    Call GetAverage_Color(tmp3, tmp2, "R1")

    ' 233.LSB定義.HL_SENR1
    Dim HL_MGERR_LSB() As Double
     HL_MGERR_LSB = HL_MGERR_DevInfo.Lsb.AsDouble

    ' 238.LSB換算.HL_SENR1
    Dim tmp4(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp4(site) = tmp3(site) * HL_MGERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.HL_SENR1
    Call ResultAdd("HL_SENR1", tmp4)

' #### HL_SENGR1 ####

    ' 0.画像情報インポート.HL_SENGR1

    ' 1.Clamp.HL_SENGR1

    ' 2.Median.HL_SENGR1

    ' 3.Median.HL_SENGR1

    ' 82.Average_FA.HL_SENGR1

    ' 83.GetAverage_Color.HL_SENGR1
    Dim tmp5(nSite) As Double
    Call GetAverage_Color(tmp5, tmp2, "Gr1")

    ' 233.LSB定義.HL_SENGR1

    ' 238.LSB換算.HL_SENGR1
    Dim tmp6(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp6(site) = tmp5(site) * HL_MGERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.HL_SENGR1
    Call ResultAdd("HL_SENGR1", tmp6)

' #### HL_SENGB1 ####

    ' 0.画像情報インポート.HL_SENGB1

    ' 1.Clamp.HL_SENGB1

    ' 2.Median.HL_SENGB1

    ' 3.Median.HL_SENGB1

    ' 82.Average_FA.HL_SENGB1

    ' 83.GetAverage_Color.HL_SENGB1
    Dim tmp7(nSite) As Double
    Call GetAverage_Color(tmp7, tmp2, "Gb1")

    ' 233.LSB定義.HL_SENGB1

    ' 238.LSB換算.HL_SENGB1
    Dim tmp8(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp8(site) = tmp7(site) * HL_MGERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.HL_SENGB1
    Call ResultAdd("HL_SENGB1", tmp8)

' #### HL_SENB1 ####

    ' 0.画像情報インポート.HL_SENB1

    ' 1.Clamp.HL_SENB1

    ' 2.Median.HL_SENB1

    ' 3.Median.HL_SENB1

    ' 82.Average_FA.HL_SENB1

    ' 83.GetAverage_Color.HL_SENB1
    Dim tmp9(nSite) As Double
    Call GetAverage_Color(tmp9, tmp2, "B1")

    ' 233.LSB定義.HL_SENB1

    ' 238.LSB換算.HL_SENB1
    Dim tmp10(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp10(site) = tmp9(site) * HL_MGERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.HL_SENB1
    Call ResultAdd("HL_SENB1", tmp10)

' #### HL_SENR2 ####

    ' 0.画像情報インポート.HL_SENR2

    ' 1.Clamp.HL_SENR2

    ' 2.Median.HL_SENR2

    ' 3.Median.HL_SENR2

    ' 82.Average_FA.HL_SENR2

    ' 83.GetAverage_Color.HL_SENR2
    Dim tmp11(nSite) As Double
    Call GetAverage_Color(tmp11, tmp2, "R2")

    ' 233.LSB定義.HL_SENR2

    ' 238.LSB換算.HL_SENR2
    Dim tmp12(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp12(site) = tmp11(site) * HL_MGERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.HL_SENR2
    Call ResultAdd("HL_SENR2", tmp12)

' #### HL_SENGR2 ####

    ' 0.画像情報インポート.HL_SENGR2

    ' 1.Clamp.HL_SENGR2

    ' 2.Median.HL_SENGR2

    ' 3.Median.HL_SENGR2

    ' 82.Average_FA.HL_SENGR2

    ' 83.GetAverage_Color.HL_SENGR2
    Dim tmp13(nSite) As Double
    Call GetAverage_Color(tmp13, tmp2, "Gr2")

    ' 233.LSB定義.HL_SENGR2

    ' 238.LSB換算.HL_SENGR2
    Dim tmp14(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp14(site) = tmp13(site) * HL_MGERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.HL_SENGR2
    Call ResultAdd("HL_SENGR2", tmp14)

' #### HL_SENGB2 ####

    ' 0.画像情報インポート.HL_SENGB2

    ' 1.Clamp.HL_SENGB2

    ' 2.Median.HL_SENGB2

    ' 3.Median.HL_SENGB2

    ' 82.Average_FA.HL_SENGB2

    ' 83.GetAverage_Color.HL_SENGB2
    Dim tmp15(nSite) As Double
    Call GetAverage_Color(tmp15, tmp2, "Gb2")

    ' 233.LSB定義.HL_SENGB2

    ' 238.LSB換算.HL_SENGB2
    Dim tmp16(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp16(site) = tmp15(site) * HL_MGERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.HL_SENGB2
    Call ResultAdd("HL_SENGB2", tmp16)

' #### HL_SENB2 ####

    ' 0.画像情報インポート.HL_SENB2

    ' 1.Clamp.HL_SENB2

    ' 2.Median.HL_SENB2

    ' 3.Median.HL_SENB2

    ' 82.Average_FA.HL_SENB2

    ' 83.GetAverage_Color.HL_SENB2
    Dim tmp17(nSite) As Double
    Call GetAverage_Color(tmp17, tmp2, "B2")

    ' 233.LSB定義.HL_SENB2

    ' 238.LSB換算.HL_SENB2
    Dim tmp18(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp18(site) = tmp17(site) * HL_MGERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.HL_SENB2
    Call ResultAdd("HL_SENB2", tmp18)

End Function



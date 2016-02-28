Attribute VB_Name = "Image_028_OFRB1_ERR_Mod"

Option Explicit

Public Function OFRB1_ERR_Process()

        Call PutImageInto_Common

' #### OFRB1_SENR ####

    Dim site As Long

    ' 0.画像情報インポート.OFRB1_SENR
    Dim OFRB1_ERR_Param As CParamPlane
    Dim OFRB1_ERR_DevInfo As CDeviceConfigInfo
    Dim OFRB1_ERR_Plane As CImgPlane
    Set OFRB1_ERR_Param = TheParameterBank.Item("OFRB1ImageTest_Acq1")
    Set OFRB1_ERR_DevInfo = TheDeviceProfiler.ConfigInfo("OFRB1ImageTest_Acq1")
    Set OFRB1_ERR_Plane = OFRB1_ERR_Param.plane

    ' 1.Clamp.OFRB1_SENR
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call Clamp(OFRB1_ERR_Plane, sPlane1, "Bayer2x4_VOPB")

    ' 2.Median.OFRB1_SENR
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 1, 5)
        Call ReleasePlane(sPlane1)

    ' 3.Median.OFRB1_SENR
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call MedianEx(sPlane2, sPlane3, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane2)

    ' 82.Average_FA.OFRB1_SENR
    Dim tmp1_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONE0", EEE_COLOR_ALL, tmp1_0)
        Call ReleasePlane(sPlane3)
    Dim tmp2 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp2, tmp1_0)

    ' 83.GetAverage_Color.OFRB1_SENR
    Dim tmp3(nSite) As Double
    Call GetAverage_Color(tmp3, tmp2, "R1", "R2")

    ' 233.LSB定義.OFRB1_SENR
    Dim OFRB1_ERR_LSB() As Double
     OFRB1_ERR_LSB = OFRB1_ERR_DevInfo.Lsb.AsDouble

    ' 238.LSB換算.OFRB1_SENR
    Dim tmp4(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp4(site) = tmp3(site) * OFRB1_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.OFRB1_SENR
    Call ResultAdd("OFRB1_SENR", tmp4)

' #### OFRB1_SENGR ####

    ' 0.画像情報インポート.OFRB1_SENGR

    ' 1.Clamp.OFRB1_SENGR

    ' 2.Median.OFRB1_SENGR

    ' 3.Median.OFRB1_SENGR

    ' 82.Average_FA.OFRB1_SENGR

    ' 83.GetAverage_Color.OFRB1_SENGR
    Dim tmp5(nSite) As Double
    Call GetAverage_Color(tmp5, tmp2, "Gr1", "Gr2")

    ' 233.LSB定義.OFRB1_SENGR

    ' 238.LSB換算.OFRB1_SENGR
    Dim tmp6(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp6(site) = tmp5(site) * OFRB1_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.OFRB1_SENGR
    Call ResultAdd("OFRB1_SENGR", tmp6)

' #### OFRB1_SENGB ####

    ' 0.画像情報インポート.OFRB1_SENGB

    ' 1.Clamp.OFRB1_SENGB

    ' 2.Median.OFRB1_SENGB

    ' 3.Median.OFRB1_SENGB

    ' 82.Average_FA.OFRB1_SENGB

    ' 83.GetAverage_Color.OFRB1_SENGB
    Dim tmp7(nSite) As Double
    Call GetAverage_Color(tmp7, tmp2, "Gb1", "Gb2")

    ' 233.LSB定義.OFRB1_SENGB

    ' 238.LSB換算.OFRB1_SENGB
    Dim tmp8(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp8(site) = tmp7(site) * OFRB1_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.OFRB1_SENGB
    Call ResultAdd("OFRB1_SENGB", tmp8)

End Function



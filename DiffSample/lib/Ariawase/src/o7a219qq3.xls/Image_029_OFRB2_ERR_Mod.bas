Attribute VB_Name = "Image_029_OFRB2_ERR_Mod"

Option Explicit

Public Function OFRB2_ERR_Process()

        Call PutImageInto_Common

' #### OFRB2_SENR ####

    Dim site As Long

    ' 0.画像情報インポート.OFRB2_SENR
    Dim OFRB2_ERR_Param As CParamPlane
    Dim OFRB2_ERR_DevInfo As CDeviceConfigInfo
    Dim OFRB2_ERR_Plane As CImgPlane
    Set OFRB2_ERR_Param = TheParameterBank.Item("OFRB2ImageTest_Acq1")
    Set OFRB2_ERR_DevInfo = TheDeviceProfiler.ConfigInfo("OFRB2ImageTest_Acq1")
    Set OFRB2_ERR_Plane = OFRB2_ERR_Param.plane

    ' 1.Clamp.OFRB2_SENR
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call Clamp(OFRB2_ERR_Plane, sPlane1, "Bayer2x4_VOPB")

    ' 2.Median.OFRB2_SENR
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 1, 5)
        Call ReleasePlane(sPlane1)

    ' 3.Median.OFRB2_SENR
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call MedianEx(sPlane2, sPlane3, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane2)

    ' 82.Average_FA.OFRB2_SENR
    Dim tmp1_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONE0", EEE_COLOR_ALL, tmp1_0)
        Call ReleasePlane(sPlane3)
    Dim tmp2 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp2, tmp1_0)

    ' 83.GetAverage_Color.OFRB2_SENR
    Dim tmp3(nSite) As Double
    Call GetAverage_Color(tmp3, tmp2, "R1", "R2")

    ' 233.LSB定義.OFRB2_SENR
    Dim OFRB2_ERR_LSB() As Double
     OFRB2_ERR_LSB = OFRB2_ERR_DevInfo.Lsb.AsDouble

    ' 238.LSB換算.OFRB2_SENR
    Dim tmp4(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp4(site) = tmp3(site) * OFRB2_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.OFRB2_SENR
    Call ResultAdd("OFRB2_SENR", tmp4)

' #### OFRB2_SENGR ####

    ' 0.画像情報インポート.OFRB2_SENGR

    ' 1.Clamp.OFRB2_SENGR

    ' 2.Median.OFRB2_SENGR

    ' 3.Median.OFRB2_SENGR

    ' 82.Average_FA.OFRB2_SENGR

    ' 83.GetAverage_Color.OFRB2_SENGR
    Dim tmp5(nSite) As Double
    Call GetAverage_Color(tmp5, tmp2, "Gr1", "Gr2")

    ' 233.LSB定義.OFRB2_SENGR

    ' 238.LSB換算.OFRB2_SENGR
    Dim tmp6(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp6(site) = tmp5(site) * OFRB2_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.OFRB2_SENGR
    Call ResultAdd("OFRB2_SENGR", tmp6)

' #### OFRB2_SENGB ####

    ' 0.画像情報インポート.OFRB2_SENGB

    ' 1.Clamp.OFRB2_SENGB

    ' 2.Median.OFRB2_SENGB

    ' 3.Median.OFRB2_SENGB

    ' 82.Average_FA.OFRB2_SENGB

    ' 83.GetAverage_Color.OFRB2_SENGB
    Dim tmp7(nSite) As Double
    Call GetAverage_Color(tmp7, tmp2, "Gb1", "Gb2")

    ' 233.LSB定義.OFRB2_SENGB

    ' 238.LSB換算.OFRB2_SENGB
    Dim tmp8(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp8(site) = tmp7(site) * OFRB2_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.OFRB2_SENGB
    Call ResultAdd("OFRB2_SENGB", tmp8)

End Function



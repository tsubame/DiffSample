Attribute VB_Name = "Image_006_DK_SGMERR_Mod"

Option Explicit

Public Function DK_SGMERR_Process1()

        Call PutImageInto_Common

End Function

Public Function DK_SGMERR_Process2()

        Call PutImageInto_Common

' #### DK_ZSGM ####

    Dim site As Long

    ' 0.複数画像情報インポート.DK_ZSGM
    Dim DK_SGMERR_0_Param As CParamPlane
    Dim DK_SGMERR_0_DevInfo As CDeviceConfigInfo
    Dim DK_SGMERR_0_Plane As CImgPlane
    Set DK_SGMERR_0_Param = TheParameterBank.Item("DK_SGMImageTest1_Acq1")
    Set DK_SGMERR_0_DevInfo = TheDeviceProfiler.ConfigInfo("DK_SGMImageTest1_Acq1")
        Call TheParameterBank.Delete("DK_SGMImageTest1_Acq1")
    Set DK_SGMERR_0_Plane = DK_SGMERR_0_Param.plane

    ' 1.複数画像情報インポート.DK_ZSGM
    Dim DK_SGMERR_1_Param As CParamPlane
    Dim DK_SGMERR_1_DevInfo As CDeviceConfigInfo
    Dim DK_SGMERR_1_Plane As CImgPlane
    Set DK_SGMERR_1_Param = TheParameterBank.Item("DK_SGMImageTest2_Acq1")
    Set DK_SGMERR_1_DevInfo = TheDeviceProfiler.ConfigInfo("DK_SGMImageTest2_Acq1")
        Call TheParameterBank.Delete("DK_SGMImageTest2_Acq1")
    Set DK_SGMERR_1_Plane = DK_SGMERR_1_Param.plane

    ' 2.Subtract(通常).DK_ZSGM
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call Subtract(DK_SGMERR_0_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, DK_SGMERR_1_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane1, "Bayer2x4_FULL", EEE_COLOR_ALL)

    ' 3.StdDev_FA.DK_ZSGM
    Dim tmp1 As CImgColorAllResult
    Call StdDev_FA(sPlane1, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, tmp1)
        Call ReleasePlane(sPlane1)

    ' 4.GetSum_Color.DK_ZSGM
    Dim tmp2(nSite) As Double
    Call GetSum_Color(tmp2, tmp1, "-")

    ' 5.複数画像用_LSB定義.DK_ZSGM
    Dim DK_SGMERR_LSB() As Double
     DK_SGMERR_LSB = DK_SGMERR_0_DevInfo.Lsb.AsDouble

    ' 6.LSB換算.DK_ZSGM
    Dim tmp3(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp3(site) = tmp2(site) * DK_SGMERR_LSB(site)
        End If
    Next site

    ' 7.PutTestResult.DK_ZSGM
    Call ResultAdd("DK_ZSGM", tmp3)

End Function



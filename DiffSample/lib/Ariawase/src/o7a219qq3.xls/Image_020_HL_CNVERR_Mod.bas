Attribute VB_Name = "Image_020_HL_CNVERR_Mod"

Option Explicit

Public Function HL_CNVERR_Process1()

        Call PutImageInto_Common

End Function

Public Function HL_CNVERR_Process2()

        Call PutImageInto_Common

End Function

Public Function HL_CNVERR_Process3()

        Call PutImageInto_Common

End Function

Public Function HL_CNVERR_Process4()

        Call PutImageInto_Common

' #### HL_CNV ####

    Dim site As Long

    ' 0.複数画像情報インポート.HL_CNV
    Dim HL_CNVERR_0_Param As CParamPlane
    Dim HL_CNVERR_0_DevInfo As CDeviceConfigInfo
    Dim HL_CNVERR_0_Plane As CImgPlane
    Set HL_CNVERR_0_Param = TheParameterBank.Item("HL_CNVImageTest1_Acq1")
    Set HL_CNVERR_0_DevInfo = TheDeviceProfiler.ConfigInfo("HL_CNVImageTest1_Acq1")
        Call TheParameterBank.Delete("HL_CNVImageTest1_Acq1")
    Set HL_CNVERR_0_Plane = HL_CNVERR_0_Param.plane

    ' 1.ShiftLeft.HL_CNV
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call ShiftLeft(HL_CNVERR_0_Plane, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane1, "Bayer2x4_ZONE22", EEE_COLOR_ALL, 2)

    ' 2.Copy.HL_CNV
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthF32, , "sPlane2")
    Call Copy(sPlane1, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane2, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane1)

    ' 3.複数画像情報インポート.HL_CNV
    Dim HL_CNVERR_1_Param As CParamPlane
    Dim HL_CNVERR_1_DevInfo As CDeviceConfigInfo
    Dim HL_CNVERR_1_Plane As CImgPlane
    Set HL_CNVERR_1_Param = TheParameterBank.Item("HL_CNVImageTest2_Acq1")
    Set HL_CNVERR_1_DevInfo = TheDeviceProfiler.ConfigInfo("HL_CNVImageTest2_Acq1")
        Call TheParameterBank.Delete("HL_CNVImageTest2_Acq1")
    Set HL_CNVERR_1_Plane = HL_CNVERR_1_Param.plane

    ' 4.Add(通常).HL_CNV
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, True, "sPlane3")
    Call Add(HL_CNVERR_0_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, HL_CNVERR_1_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane3, "Bayer2x4_FULL", EEE_COLOR_ALL)

    ' 5.ShiftLeft.HL_CNV
    Dim sPlane4 As CImgPlane
    Call GetFreePlane(sPlane4, "Normal_Bayer2x4", idpDepthS16, , "sPlane4")
    Call ShiftLeft(HL_CNVERR_1_Plane, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane4, "Bayer2x4_ZONE22", EEE_COLOR_ALL, 2)

    ' 6.Copy.HL_CNV
    Dim sPlane5 As CImgPlane
    Call GetFreePlane(sPlane5, "Normal_Bayer2x4", idpDepthF32, , "sPlane5")
    Call Copy(sPlane4, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane5, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane4)

    ' 7.複数画像情報インポート.HL_CNV
    Dim HL_CNVERR_2_Param As CParamPlane
    Dim HL_CNVERR_2_DevInfo As CDeviceConfigInfo
    Dim HL_CNVERR_2_Plane As CImgPlane
    Set HL_CNVERR_2_Param = TheParameterBank.Item("HL_CNVImageTest3_Acq1")
    Set HL_CNVERR_2_DevInfo = TheDeviceProfiler.ConfigInfo("HL_CNVImageTest3_Acq1")
        Call TheParameterBank.Delete("HL_CNVImageTest3_Acq1")
    Set HL_CNVERR_2_Plane = HL_CNVERR_2_Param.plane

    ' 8.ShiftLeft.HL_CNV
    Dim sPlane6 As CImgPlane
    Call GetFreePlane(sPlane6, "Normal_Bayer2x4", idpDepthS16, , "sPlane6")
    Call ShiftLeft(HL_CNVERR_2_Plane, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane6, "Bayer2x4_ZONE22", EEE_COLOR_ALL, 2)

    ' 9.Copy.HL_CNV
    Dim sPlane7 As CImgPlane
    Call GetFreePlane(sPlane7, "Normal_Bayer2x4", idpDepthF32, , "sPlane7")
    Call Copy(sPlane6, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane7, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane6)

    ' 10.Add(通常).HL_CNV
    Dim sPlane8 As CImgPlane
    Call GetFreePlane(sPlane8, "Normal_Bayer2x4", idpDepthS16, True, "sPlane8")
    Call Add(sPlane3, "Bayer2x4_FULL", EEE_COLOR_ALL, HL_CNVERR_2_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane8, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane3)

    ' 11.複数画像情報インポート.HL_CNV
    Dim HL_CNVERR_3_Param As CParamPlane
    Dim HL_CNVERR_3_DevInfo As CDeviceConfigInfo
    Dim HL_CNVERR_3_Plane As CImgPlane
    Set HL_CNVERR_3_Param = TheParameterBank.Item("HL_CNVImageTest4_Acq1")
    Set HL_CNVERR_3_DevInfo = TheDeviceProfiler.ConfigInfo("HL_CNVImageTest4_Acq1")
        Call TheParameterBank.Delete("HL_CNVImageTest4_Acq1")
    Set HL_CNVERR_3_Plane = HL_CNVERR_3_Param.plane

    ' 12.ShiftLeft.HL_CNV
    Dim sPlane9 As CImgPlane
    Call GetFreePlane(sPlane9, "Normal_Bayer2x4", idpDepthS16, , "sPlane9")
    Call ShiftLeft(HL_CNVERR_3_Plane, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane9, "Bayer2x4_ZONE22", EEE_COLOR_ALL, 2)

    ' 13.Copy.HL_CNV
    Dim sPlane10 As CImgPlane
    Call GetFreePlane(sPlane10, "Normal_Bayer2x4", idpDepthF32, , "sPlane10")
    Call Copy(sPlane9, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane10, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane9)

    ' 14.Add(通常).HL_CNV
    Dim sPlane11 As CImgPlane
    Call GetFreePlane(sPlane11, "Normal_Bayer2x4", idpDepthS16, True, "sPlane11")
    Call Add(sPlane8, "Bayer2x4_FULL", EEE_COLOR_ALL, HL_CNVERR_3_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane11, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane8)

    ' 15.Copy.HL_CNV
    Dim sPlane12 As CImgPlane
    Call GetFreePlane(sPlane12, "Normal_Bayer2x4", idpDepthF32, , "sPlane12")
    Call Copy(sPlane11, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane12, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane11)

    ' 16.Subtract(通常).HL_CNV
    Dim sPlane13 As CImgPlane
    Call GetFreePlane(sPlane13, "Normal_Bayer2x4", idpDepthF32, , "sPlane13")
    Call Subtract(sPlane7, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane12, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane13, "Bayer2x4_ZONE22", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane7)

    ' 17.Multiply(通常).HL_CNV
    Dim sPlane14 As CImgPlane
    Call GetFreePlane(sPlane14, "Normal_Bayer2x4", idpDepthF32, , "sPlane14")
    Call Multiply(sPlane13, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane13, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane14, "Bayer2x4_ZONE22", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane13)

    ' 18.Subtract(通常).HL_CNV
    Dim sPlane15 As CImgPlane
    Call GetFreePlane(sPlane15, "Normal_Bayer2x4", idpDepthF32, , "sPlane15")
    Call Subtract(sPlane10, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane12, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane15, "Bayer2x4_ZONE22", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane10)

    ' 19.Multiply(通常).HL_CNV
    Dim sPlane16 As CImgPlane
    Call GetFreePlane(sPlane16, "Normal_Bayer2x4", idpDepthF32, , "sPlane16")
    Call Multiply(sPlane15, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane15, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane16, "Bayer2x4_ZONE22", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane15)

    ' 20.Subtract(通常).HL_CNV
    Dim sPlane17 As CImgPlane
    Call GetFreePlane(sPlane17, "Normal_Bayer2x4", idpDepthF32, , "sPlane17")
    Call Subtract(sPlane2, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane12, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane17, "Bayer2x4_ZONE22", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane2)

    ' 21.Multiply(通常).HL_CNV
    Dim sPlane18 As CImgPlane
    Call GetFreePlane(sPlane18, "Normal_Bayer2x4", idpDepthF32, , "sPlane18")
    Call Multiply(sPlane17, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane17, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane18, "Bayer2x4_ZONE22", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane17)

    ' 22.Subtract(通常).HL_CNV
    Dim sPlane19 As CImgPlane
    Call GetFreePlane(sPlane19, "Normal_Bayer2x4", idpDepthF32, , "sPlane19")
    Call Subtract(sPlane5, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane12, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane19, "Bayer2x4_ZONE22", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane5)

    ' 23.Multiply(通常).HL_CNV
    Dim sPlane20 As CImgPlane
    Call GetFreePlane(sPlane20, "Normal_Bayer2x4", idpDepthF32, , "sPlane20")
    Call Multiply(sPlane19, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane19, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane20, "Bayer2x4_ZONE22", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane19)

    ' 24.Add(通常).HL_CNV
    Dim sPlane21 As CImgPlane
    Call GetFreePlane(sPlane21, "Normal_Bayer2x4", idpDepthF32, True, "sPlane21")
    Call Add(sPlane18, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane20, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane21, "Bayer2x4_ZONE22", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane18)
        Call ReleasePlane(sPlane20)

    ' 25.Add(通常).HL_CNV
    Dim sPlane22 As CImgPlane
    Call GetFreePlane(sPlane22, "Normal_Bayer2x4", idpDepthF32, True, "sPlane22")
    Call Add(sPlane21, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane14, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane22, "Bayer2x4_ZONE22", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane14)
        Call ReleasePlane(sPlane21)

    ' 26.Add(通常).HL_CNV
    Dim sPlane23 As CImgPlane
    Call GetFreePlane(sPlane23, "Normal_Bayer2x4", idpDepthF32, True, "sPlane23")
    Call Add(sPlane22, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane16, "Bayer2x4_ZONE22", EEE_COLOR_ALL, sPlane23, "Bayer2x4_ZONE22", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane16)
        Call ReleasePlane(sPlane22)

    ' 28.Average_FA.HL_CNV
    Dim tmp1_0 As CImgColorAllResult
    Call Average_FA(sPlane23, "Bayer2x4_ZONE22", EEE_COLOR_ALL, tmp1_0)
        Call ReleasePlane(sPlane23)
    Dim tmp2 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp2, tmp1_0)

    ' 29.Clamp.HL_CNV
    Dim sPlane24 As CImgPlane
    Call GetFreePlane(sPlane24, "Normal_Bayer2x4", idpDepthF32, , "sPlane24")
    Call Clamp(sPlane12, sPlane24, "Bayer2x4_VOPB")
        Call ReleasePlane(sPlane12)

    ' 31.Average_FA.HL_CNV
    Dim tmp3_0 As CImgColorAllResult
    Call Average_FA(sPlane24, "Bayer2x4_ZONE22", EEE_COLOR_ALL, tmp3_0)
        Call ReleasePlane(sPlane24)
    Dim tmp4 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp4, tmp3_0)

    ' 32.複数画像用_LSB定義.HL_CNV
    Dim HL_CNVERR_LSB() As Double
     HL_CNVERR_LSB = HL_CNVERR_0_DevInfo.Lsb.AsDouble

    ' 33.計算式評価.HL_CNV
    Dim tmp5(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp5(site) = 4
        End If
    Next site

    ' 34.計算式評価.HL_CNV
    Dim tmp6(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp6(site) = tmp5(site) * tmp5(site) * (tmp5(site) - 1)
        End If
    Next site

    ' 35.AllAccValue.HL_CNV
    Dim tmp7 As CImgColorAllResult
    Call CImgColorAllResultAcc(tmp7, tmp2, "/", tmp6)

    ' 36.GetSum_Color.HL_CNV
    Dim tmp8(nSite) As Double
    Call GetSum_Color(tmp8, tmp7, "R1", "R2", "Gr1", "Gr2", "Gb1", "Gb2", "B1", "B2")

    ' 37.AllAccValue.HL_CNV
    Dim tmp9 As CImgColorAllResult
    Call CImgColorAllResultAcc(tmp9, tmp4, "/", tmp5)

    ' 38.AllAccArray.HL_CNV
    Dim tmp10 As CImgColorAllResult
    Call CImgColorAllResultAcc(tmp10, tmp9, "*", tmp7)

    ' 39.GetSum_Color.HL_CNV
    Dim tmp11(nSite) As Double
    Call GetSum_Color(tmp11, tmp10, "R1", "R2", "Gr1", "Gr2", "Gb1", "Gb2", "B1", "B2")

    ' 40.使用色情報カウント.HL_CNV
    Dim tmp12(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp12(site) = 7 + 1
        End If
    Next site

    ' 41.AllAccArray.HL_CNV
    Dim tmp13 As CImgColorAllResult
    Call CImgColorAllResultAcc(tmp13, tmp9, "*", tmp9)

    ' 42.GetSum_Color.HL_CNV
    Dim tmp14(nSite) As Double
    Call GetSum_Color(tmp14, tmp13, "R1", "R2", "Gr1", "Gr2", "Gb1", "Gb2", "B1", "B2")

    ' 43.GetSum_Color.HL_CNV
    Dim tmp15(nSite) As Double
    Call GetSum_Color(tmp15, tmp9, "R1", "R2", "Gr1", "Gr2", "Gb1", "Gb2", "B1", "B2")

    ' 44.計算式評価.HL_CNV
    Dim tmp16(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp16(site) = tmp11(site) * tmp12(site) - tmp15(site) * tmp8(site)
        End If
    Next site

    ' 45.計算式評価.HL_CNV
    Dim tmp17(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp17(site) = tmp14(site) * tmp12(site) - tmp15(site) * tmp15(site)
        End If
    Next site

    ' 46.計算式評価.HL_CNV
    Dim tmp18(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp18(site) = Div(tmp16(site), tmp17(site), 999) * HL_CNVERR_LSB(site)
        End If
    Next site

    ' 47.PutTestResult.HL_CNV
    Call ResultAdd("HL_CNV", tmp18)

End Function



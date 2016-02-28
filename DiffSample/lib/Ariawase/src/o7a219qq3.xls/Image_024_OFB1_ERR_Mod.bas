Attribute VB_Name = "Image_024_OFB1_ERR_Mod"

Option Explicit

Public Function OFB1_ERR_Process()

        Call PutImageInto_Common

' #### OFB1_QSAV ####

    Dim site As Long

    ' 0.画像情報インポート.OFB1_QSAV
    Dim OFB1_ERR_Param As CParamPlane
    Dim OFB1_ERR_DevInfo As CDeviceConfigInfo
    Dim OFB1_ERR_Plane As CImgPlane
    Set OFB1_ERR_Param = TheParameterBank.Item("OFB1ImageTest_Acq1")
    Set OFB1_ERR_DevInfo = TheDeviceProfiler.ConfigInfo("OFB1ImageTest_Acq1")
    Set OFB1_ERR_Plane = OFB1_ERR_Param.plane

    ' 1.Clamp.OFB1_QSAV
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call Clamp(OFB1_ERR_Plane, sPlane1, "Bayer2x4_VOPB")

    ' 2.Median.OFB1_QSAV
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 1, 5)

    ' 3.Median.OFB1_QSAV
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call MedianEx(sPlane2, sPlane3, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane2)

    ' 82.Average_FA.OFB1_QSAV
    Dim tmp1_0 As CImgColorAllResult
    Call Average_FA(sPlane3, "Bayer2x4_ZONE2D", EEE_COLOR_ALL, tmp1_0)
    Dim tmp2 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp2, tmp1_0)

    ' 83.GetAverage_Color.OFB1_QSAV
    Dim tmp3(nSite) As Double
    Call GetAverage_Color(tmp3, tmp2, "-")

    ' 233.LSB定義.OFB1_QSAV
    Dim OFB1_ERR_LSB() As Double
     OFB1_ERR_LSB = OFB1_ERR_DevInfo.Lsb.AsDouble

    ' 238.LSB換算.OFB1_QSAV
    Dim tmp4(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp4(site) = tmp3(site) * OFB1_ERR_LSB(site)
        End If
    Next site

    ' 239.PutTestResult.OFB1_QSAV
    Call ResultAdd("OFB1_QSAV", tmp4)

' #### OFB1_4HLN ####

    ' 0.画像情報インポート.OFB1_4HLN

    ' 1.Clamp.OFB1_4HLN

    ' 2.Median.OFB1_4HLN
    Dim sPlane4 As CImgPlane
    Call GetFreePlane(sPlane4, "Normal_Bayer2x4", idpDepthS16, , "sPlane4")
    Call MedianEx(sPlane1, sPlane4, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane1)

    ' 3.ZONE取得.OFB1_4HLN

    ' 5.AccumulateRow.OFB1_4HLN
    Dim sPlane5 As CImgPlane
    Call GetFreePlane(sPlane5, "Normal_Bayer2x4_ACR", idpDepthF32, , "sPlane5")
    Call MakeAcrPMD(sPlane5, "Bayer2x4_ZONE2D", "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT)
    Call AccumulateRow(sPlane4, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane5, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, idpAccumMean)

    ' 6.SubRows.OFB1_4HLN
    Dim sPlane6 As CImgPlane
    Call GetFreePlane(sPlane6, "Normal_Bayer2x4_ACR", idpDepthF32, , "sPlane6")
    Call SubRows(sPlane5, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, sPlane6, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, 4)
        Call ReleasePlane(sPlane5)
    Call MakeAcrJudgePMD(sPlane6, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", "Bayer2x4_ACR_4_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, 4)

    ' 9.AbsMax_FA.OFB1_4HLN
    Dim tmp5_0 As CImgColorAllResult
    Call AbsMax_FA(sPlane6, "Bayer2x4_ACR_4_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_ALL, tmp5_0)
        Call ReleasePlane(sPlane6)
    Dim tmp6 As CImgColorAllResult
    Call GetMax_CImgColor(tmp6, tmp5_0)

    ' 10.GetAbsMax_Color.OFB1_4HLN
    Dim tmp7(nSite) As Double
    Call GetAbsMax_Color(tmp7, tmp6, "-")

    ' 13.GetAbs.OFB1_4HLN
    Dim tmp8(nSite) As Double
    Call GetAbs(tmp8, tmp7)

    ' 14.パラメータ取得.OFB1_4HLN
    Dim tmp_OFB1_QSAV() As Double
    TheResult.GetResult "OFB1_QSAV", tmp_OFB1_QSAV
    OFB1_ERR_LSB = TheDeviceProfiler.ConfigInfo("OFB1ImageTest_Acq1").Lsb.AsDouble
        Call TheParameterBank.Delete("OFB1ImageTest_Acq1")
    Dim tmp9(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp9(site) = Div(tmp_OFB1_QSAV(site), OFB1_ERR_LSB(site), 0)
        End If
    Next site

    ' 15.計算式評価.OFB1_4HLN
    Dim tmp10(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp10(site) = Div(tmp8(site), tmp9(site), 999)
        End If
    Next site

    ' 17.LSB定義.OFB1_4HLN

    ' 19.PutTestResult.OFB1_4HLN
    Call ResultAdd("OFB1_4HLN", tmp10)

' #### OFB1_HOBI ####

    ' 0.画像情報インポート.OFB1_HOBI

    ' 1.Clamp.OFB1_HOBI

    ' 2.Median.OFB1_HOBI

    ' 3.Median.OFB1_HOBI

    ' 4.Median.OFB1_HOBI

    ' 5.Subtract(通常).OFB1_HOBI
    Dim sPlane7 As CImgPlane
    Call GetFreePlane(sPlane7, "Normal_Bayer2x4", idpDepthS16, , "sPlane7")
    Call Subtract(sPlane4, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane3, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane7, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane3)
        Call ReleasePlane(sPlane4)

    ' 7.ZONE取得.OFB1_HOBI

    ' 9.AccumulateRow.OFB1_HOBI
    Dim sPlane8 As CImgPlane
    Call GetFreePlane(sPlane8, "Normal_Bayer2x4_ACR", idpDepthF32, , "sPlane8")
    Call MakeAcrPMD(sPlane8, "Bayer2x4_ZONE2D", "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT)
    Call AccumulateRow(sPlane7, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, sPlane8, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_FLAT, idpAccumMean)

    ' 10.PutFlag_FA.OFB1_HOBI
    Call PutFlag_FA(sPlane8, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_ALL, idpCountBelow, 0, 0, idpLimitEachSite, idpLimitExclude, "Flg_Temp1")

    ' 11.MultiplyConstFlag.OFB1_HOBI
    Call MultiplyConstFlag(sPlane8, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_ALL, -1, sPlane8, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_ALL, "Flg_Temp1")

    ' 12.Average_FA.OFB1_HOBI
    Dim tmp11_0 As CImgColorAllResult
    Call Average_FA(sPlane8, "Bayer2x4_ACR_ZONE2D_EEE_COLOR_FLAT", EEE_COLOR_ALL, tmp11_0)
        Call ReleasePlane(sPlane8)
    Dim tmp12 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp12, tmp11_0)

    ' 13.GetAverage_Color.OFB1_HOBI
    Dim tmp13(nSite) As Double
    Call GetAverage_Color(tmp13, tmp12, "-")

    ' 14.LSB定義.OFB1_HOBI

    ' 15.LSB換算.OFB1_HOBI
    Dim tmp14(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp14(site) = tmp13(site) * OFB1_ERR_LSB(site)
        End If
    Next site

    ' 16.PutTestResult.OFB1_HOBI
    Call ResultAdd("OFB1_HOBI", tmp14)

' #### OFB1_HOBIR ####

    ' 0.画像情報インポート.OFB1_HOBIR

    ' 1.Clamp.OFB1_HOBIR

    ' 2.Median.OFB1_HOBIR

    ' 3.Median.OFB1_HOBIR

    ' 4.Median.OFB1_HOBIR

    ' 5.Subtract(通常).OFB1_HOBIR

    ' 6.カラーマップ遷移.OFB1_HOBIR
    Dim sPlane9 As CImgPlane
    Call GetFreePlane(sPlane9, "Normal_Bayer2x2", idpDepthS16, , "sPlane9")
    Call Copy(sPlane7, sPlane7.BasePMD.Name, EEE_COLOR_FLAT, sPlane9, sPlane9.BasePMD.Name, EEE_COLOR_FLAT)
        Call ReleasePlane(sPlane7)

    ' 7.ZONE取得.OFB1_HOBIR

    ' 9.AccumulateRow.OFB1_HOBIR
    Dim sPlane10 As CImgPlane
    Call GetFreePlane(sPlane10, "Normal_Bayer2x2_ACR", idpDepthF32, , "sPlane10")
    Call MakeAcrPMD(sPlane10, "Bayer2x2_ZONE2D", "Bayer2x2_ACR_ZONE2D_EEE_COLOR_ALL", EEE_COLOR_ALL)
    Call AccumulateRow(sPlane9, "Bayer2x2_ZONE2D", EEE_COLOR_ALL, sPlane10, "Bayer2x2_ACR_ZONE2D_EEE_COLOR_ALL", EEE_COLOR_ALL, idpAccumMean)
        Call ReleasePlane(sPlane9)

    ' 10.PutFlag_FA.OFB1_HOBIR
    Call PutFlag_FA(sPlane10, "Bayer2x2_ACR_ZONE2D_EEE_COLOR_ALL", EEE_COLOR_ALL, idpCountBelow, 0, 0, idpLimitEachSite, idpLimitExclude, "Flg_Temp1")
'    Call FlagColorExtract(sPlane10, "Flg_Temp1", "R")

    ' 11.MultiplyConstFlag.OFB1_HOBIR
    Call MultiplyConstFlag(sPlane10, "Bayer2x2_ACR_ZONE2D_EEE_COLOR_ALL", EEE_COLOR_ALL, -1, sPlane10, "Bayer2x2_ACR_ZONE2D_EEE_COLOR_ALL", EEE_COLOR_ALL, "Flg_Temp1")

    ' 12.Average_FA.OFB1_HOBIR
    Dim tmp15_0 As CImgColorAllResult
    Call Average_FA(sPlane10, "Bayer2x2_ACR_ZONE2D_EEE_COLOR_ALL", EEE_COLOR_ALL, tmp15_0)
    Dim tmp16 As CImgColorAllResult
    Call GetAverage_CImgColor(tmp16, tmp15_0)

    ' 13.GetAverage_Color.OFB1_HOBIR
    Dim tmp17(nSite) As Double
    Call GetAverage_Color(tmp17, tmp16, "R")

    ' 14.LSB定義.OFB1_HOBIR

    ' 15.LSB換算.OFB1_HOBIR
    Dim tmp18(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp18(site) = tmp17(site) * OFB1_ERR_LSB(site)
        End If
    Next site

    ' 16.PutTestResult.OFB1_HOBIR
    Call ResultAdd("OFB1_HOBIR", tmp18)

' #### OFB1_HOBIGR ####

    ' 0.画像情報インポート.OFB1_HOBIGR

    ' 1.Clamp.OFB1_HOBIGR

    ' 2.Median.OFB1_HOBIGR

    ' 3.Median.OFB1_HOBIGR

    ' 4.Median.OFB1_HOBIGR

    ' 5.Subtract(通常).OFB1_HOBIGR

    ' 6.カラーマップ遷移.OFB1_HOBIGR

    ' 7.ZONE取得.OFB1_HOBIGR

    ' 9.AccumulateRow.OFB1_HOBIGR

    ' 10.PutFlag_FA.OFB1_HOBIGR
'    Call FlagColorExtract(sPlane10, "Flg_Temp1", "Gr")

    ' 11.MultiplyConstFlag.OFB1_HOBIGR

    ' 12.Average_FA.OFB1_HOBIGR

    ' 13.GetAverage_Color.OFB1_HOBIGR
    Dim tmp19(nSite) As Double
    Call GetAverage_Color(tmp19, tmp16, "Gr")

    ' 14.LSB定義.OFB1_HOBIGR

    ' 15.LSB換算.OFB1_HOBIGR
    Dim tmp20(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp20(site) = tmp19(site) * OFB1_ERR_LSB(site)
        End If
    Next site

    ' 16.PutTestResult.OFB1_HOBIGR
    Call ResultAdd("OFB1_HOBIGR", tmp20)

' #### OFB1_HOBIGB ####

    ' 0.画像情報インポート.OFB1_HOBIGB

    ' 1.Clamp.OFB1_HOBIGB

    ' 2.Median.OFB1_HOBIGB

    ' 3.Median.OFB1_HOBIGB

    ' 4.Median.OFB1_HOBIGB

    ' 5.Subtract(通常).OFB1_HOBIGB

    ' 6.カラーマップ遷移.OFB1_HOBIGB

    ' 7.ZONE取得.OFB1_HOBIGB

    ' 9.AccumulateRow.OFB1_HOBIGB

    ' 10.PutFlag_FA.OFB1_HOBIGB
'    Call FlagColorExtract(sPlane10, "Flg_Temp1", "Gb")

    ' 11.MultiplyConstFlag.OFB1_HOBIGB

    ' 12.Average_FA.OFB1_HOBIGB

    ' 13.GetAverage_Color.OFB1_HOBIGB
    Dim tmp21(nSite) As Double
    Call GetAverage_Color(tmp21, tmp16, "Gb")

    ' 14.LSB定義.OFB1_HOBIGB

    ' 15.LSB換算.OFB1_HOBIGB
    Dim tmp22(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp22(site) = tmp21(site) * OFB1_ERR_LSB(site)
        End If
    Next site

    ' 16.PutTestResult.OFB1_HOBIGB
    Call ResultAdd("OFB1_HOBIGB", tmp22)

' #### OFB1_HOBIB ####

    ' 0.画像情報インポート.OFB1_HOBIB

    ' 1.Clamp.OFB1_HOBIB

    ' 2.Median.OFB1_HOBIB

    ' 3.Median.OFB1_HOBIB

    ' 4.Median.OFB1_HOBIB

    ' 5.Subtract(通常).OFB1_HOBIB

    ' 6.カラーマップ遷移.OFB1_HOBIB

    ' 7.ZONE取得.OFB1_HOBIB

    ' 9.AccumulateRow.OFB1_HOBIB

    ' 10.PutFlag_FA.OFB1_HOBIB
'    Call FlagColorExtract(sPlane10, "Flg_Temp1", "B")
        Call ClearALLFlagBit("Flg_Temp1")
        Call ReleasePlane(sPlane10)

    ' 11.MultiplyConstFlag.OFB1_HOBIB

    ' 12.Average_FA.OFB1_HOBIB

    ' 13.GetAverage_Color.OFB1_HOBIB
    Dim tmp23(nSite) As Double
    Call GetAverage_Color(tmp23, tmp16, "B")

    ' 14.LSB定義.OFB1_HOBIB

    ' 15.LSB換算.OFB1_HOBIB
    Dim tmp24(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tmp24(site) = tmp23(site) * OFB1_ERR_LSB(site)
        End If
    Next site

    ' 16.PutTestResult.OFB1_HOBIB
    Call ResultAdd("OFB1_HOBIB", tmp24)

End Function



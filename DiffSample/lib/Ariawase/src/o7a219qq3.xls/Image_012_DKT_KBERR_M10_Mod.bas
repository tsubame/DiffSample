Attribute VB_Name = "Image_012_DKT_KBERR_M10_Mod"

Option Explicit

Public Function DKT_KBERR_M10_Process()

        Call PutImageInto_Common

' #### DKT_KBV001_M10 ####

    Dim site As Long

    ' 0.画像情報インポート.DKT_KBV001_M10
    Dim DKT_KBERR_M10_Param As CParamPlane
    Dim DKT_KBERR_M10_DevInfo As CDeviceConfigInfo
    Dim DKT_KBERR_M10_Plane As CImgPlane
    Set DKT_KBERR_M10_Param = TheParameterBank.Item("DKT_KB_M10ImageTest_Acq1")
    Set DKT_KBERR_M10_DevInfo = TheDeviceProfiler.ConfigInfo("DKT_KB_M10ImageTest_Acq1")
        Call TheParameterBank.Delete("DKT_KB_M10ImageTest_Acq1")
    Set DKT_KBERR_M10_Plane = DKT_KBERR_M10_Param.plane

    ' 2.Median.DKT_KBV001_M10
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call MedianEx(DKT_KBERR_M10_Plane, sPlane1, "Bayer2x4_ZONE3", 1, 5)

    ' 3.Median.DKT_KBV001_M10
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane1)

    ' 4.Subtract(通常).DKT_KBV001_M10
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call Subtract(DKT_KBERR_M10_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane2, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane3, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane2)

    ' 5.LSB定義.DKT_KBV001_M10
    Dim DKT_KBERR_M10_LSB() As Double
     DKT_KBERR_M10_LSB = DKT_KBERR_M10_DevInfo.Lsb.AsDouble

    ' 6.SliceLevel生成_ShiroKobu_マージプロセス.DKT_KBV001_M10
    Dim tmp_Slice1(nSite, (0.01 - 0.0001) / 0.0001) As Double
    Call MakeSliceLevel_ShiroKobu_1Step(tmp_Slice1, 0.0001, 0.01, 0.0001, DKT_KBERR_M10_LSB)

    ' 7.Count_ShiroKobu_マージプロセス.DKT_KBV001_M10
    Dim tmp1(nSite, (0.01 - 0.0001) / 0.0001) As Double
    Call Count_ShiroKobu_marge(sPlane3, "Bayer2x4_ZONE2D", tmp_Slice1, tmp1, "Between", "Above")
        Call ReleasePlane(sPlane3)

    ' 8.PutTestResult_マージプロセス.DKT_KBV001_M10
    Call ResultAdd_ShiroKobu("DKT_KBV001_M10", tmp1, 1 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV002_M10", tmp1, 1 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV003_M10", tmp1, 1 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV004_M10", tmp1, 1 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV005_M10", tmp1, 1 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV006_M10", tmp1, 1 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV007_M10", tmp1, 1 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV008_M10", tmp1, 1 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV009_M10", tmp1, 1 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV010_M10", tmp1, 1 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV011_M10", tmp1, 2 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV012_M10", tmp1, 2 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV013_M10", tmp1, 2 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV014_M10", tmp1, 2 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV015_M10", tmp1, 2 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV016_M10", tmp1, 2 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV017_M10", tmp1, 2 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV018_M10", tmp1, 2 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV019_M10", tmp1, 2 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV020_M10", tmp1, 2 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV021_M10", tmp1, 3 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV022_M10", tmp1, 3 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV023_M10", tmp1, 3 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV024_M10", tmp1, 3 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV025_M10", tmp1, 3 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV026_M10", tmp1, 3 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV027_M10", tmp1, 3 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV028_M10", tmp1, 3 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV029_M10", tmp1, 3 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV030_M10", tmp1, 3 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV031_M10", tmp1, 4 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV032_M10", tmp1, 4 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV033_M10", tmp1, 4 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV034_M10", tmp1, 4 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV035_M10", tmp1, 4 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV036_M10", tmp1, 4 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV037_M10", tmp1, 4 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV038_M10", tmp1, 4 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV039_M10", tmp1, 4 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV040_M10", tmp1, 4 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV041_M10", tmp1, 5 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV042_M10", tmp1, 5 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV043_M10", tmp1, 5 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV044_M10", tmp1, 5 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV045_M10", tmp1, 5 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV046_M10", tmp1, 5 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV047_M10", tmp1, 5 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV048_M10", tmp1, 5 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV049_M10", tmp1, 5 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV050_M10", tmp1, 5 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV051_M10", tmp1, 6 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV052_M10", tmp1, 6 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV053_M10", tmp1, 6 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV054_M10", tmp1, 6 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV055_M10", tmp1, 6 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV056_M10", tmp1, 6 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV057_M10", tmp1, 6 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV058_M10", tmp1, 6 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV059_M10", tmp1, 6 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV060_M10", tmp1, 6 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV061_M10", tmp1, 7 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV062_M10", tmp1, 7 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV063_M10", tmp1, 7 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV064_M10", tmp1, 7 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV065_M10", tmp1, 7 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV066_M10", tmp1, 7 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV067_M10", tmp1, 7 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV068_M10", tmp1, 7 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV069_M10", tmp1, 7 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV070_M10", tmp1, 7 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV071_M10", tmp1, 8 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV072_M10", tmp1, 8 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV073_M10", tmp1, 8 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV074_M10", tmp1, 8 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV075_M10", tmp1, 8 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV076_M10", tmp1, 8 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV077_M10", tmp1, 8 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV078_M10", tmp1, 8 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV079_M10", tmp1, 8 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV080_M10", tmp1, 8 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV081_M10", tmp1, 9 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV082_M10", tmp1, 9 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV083_M10", tmp1, 9 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV084_M10", tmp1, 9 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV085_M10", tmp1, 9 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV086_M10", tmp1, 9 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV087_M10", tmp1, 9 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV088_M10", tmp1, 9 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV089_M10", tmp1, 9 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV090_M10", tmp1, 9 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV091_M10", tmp1, 10 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV092_M10", tmp1, 10 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV093_M10", tmp1, 10 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV094_M10", tmp1, 10 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV095_M10", tmp1, 10 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV096_M10", tmp1, 10 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV097_M10", tmp1, 10 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV098_M10", tmp1, 10 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV099_M10", tmp1, 10 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV100_M10", tmp1, 10 * 10 + 9 - 10)

End Function



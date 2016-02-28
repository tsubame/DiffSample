Attribute VB_Name = "Image_013_DKT_KBERR_M16_Mod"

Option Explicit

Public Function DKT_KBERR_M16_Process()

        Call PutImageInto_Common

' #### DKT_KBV001_M16 ####

    Dim site As Long

    ' 0.画像情報インポート.DKT_KBV001_M16
    Dim DKT_KBERR_M16_Param As CParamPlane
    Dim DKT_KBERR_M16_DevInfo As CDeviceConfigInfo
    Dim DKT_KBERR_M16_Plane As CImgPlane
    Set DKT_KBERR_M16_Param = TheParameterBank.Item("DKT_KB_M16ImageTest_Acq1")
    Set DKT_KBERR_M16_DevInfo = TheDeviceProfiler.ConfigInfo("DKT_KB_M16ImageTest_Acq1")
        Call TheParameterBank.Delete("DKT_KB_M16ImageTest_Acq1")
    Set DKT_KBERR_M16_Plane = DKT_KBERR_M16_Param.plane

    ' 2.Median.DKT_KBV001_M16
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call MedianEx(DKT_KBERR_M16_Plane, sPlane1, "Bayer2x4_ZONE3", 1, 5)

    ' 3.Median.DKT_KBV001_M16
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane1)

    ' 4.Subtract(通常).DKT_KBV001_M16
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call Subtract(DKT_KBERR_M16_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane2, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane3, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane2)

    ' 5.LSB定義.DKT_KBV001_M16
    Dim DKT_KBERR_M16_LSB() As Double
     DKT_KBERR_M16_LSB = DKT_KBERR_M16_DevInfo.Lsb.AsDouble

    ' 6.SliceLevel生成_ShiroKobu_マージプロセス.DKT_KBV001_M16
    Dim tmp_Slice1(nSite, (0.01 - 0.0001) / 0.0001) As Double
    Call MakeSliceLevel_ShiroKobu_1Step(tmp_Slice1, 0.0001, 0.01, 0.0001, DKT_KBERR_M16_LSB)

    ' 7.Count_ShiroKobu_マージプロセス.DKT_KBV001_M16
    Dim tmp1(nSite, (0.01 - 0.0001) / 0.0001) As Double
    Call Count_ShiroKobu_marge(sPlane3, "Bayer2x4_ZONE2D", tmp_Slice1, tmp1, "Between", "Above")
        Call ReleasePlane(sPlane3)

    ' 8.PutTestResult_マージプロセス.DKT_KBV001_M16
    Call ResultAdd_ShiroKobu("DKT_KBV001_M16", tmp1, 1 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV002_M16", tmp1, 1 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV003_M16", tmp1, 1 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV004_M16", tmp1, 1 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV005_M16", tmp1, 1 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV006_M16", tmp1, 1 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV007_M16", tmp1, 1 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV008_M16", tmp1, 1 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV009_M16", tmp1, 1 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV010_M16", tmp1, 1 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV011_M16", tmp1, 2 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV012_M16", tmp1, 2 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV013_M16", tmp1, 2 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV014_M16", tmp1, 2 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV015_M16", tmp1, 2 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV016_M16", tmp1, 2 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV017_M16", tmp1, 2 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV018_M16", tmp1, 2 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV019_M16", tmp1, 2 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV020_M16", tmp1, 2 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV021_M16", tmp1, 3 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV022_M16", tmp1, 3 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV023_M16", tmp1, 3 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV024_M16", tmp1, 3 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV025_M16", tmp1, 3 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV026_M16", tmp1, 3 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV027_M16", tmp1, 3 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV028_M16", tmp1, 3 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV029_M16", tmp1, 3 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV030_M16", tmp1, 3 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV031_M16", tmp1, 4 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV032_M16", tmp1, 4 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV033_M16", tmp1, 4 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV034_M16", tmp1, 4 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV035_M16", tmp1, 4 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV036_M16", tmp1, 4 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV037_M16", tmp1, 4 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV038_M16", tmp1, 4 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV039_M16", tmp1, 4 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV040_M16", tmp1, 4 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV041_M16", tmp1, 5 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV042_M16", tmp1, 5 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV043_M16", tmp1, 5 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV044_M16", tmp1, 5 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV045_M16", tmp1, 5 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV046_M16", tmp1, 5 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV047_M16", tmp1, 5 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV048_M16", tmp1, 5 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV049_M16", tmp1, 5 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV050_M16", tmp1, 5 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV051_M16", tmp1, 6 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV052_M16", tmp1, 6 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV053_M16", tmp1, 6 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV054_M16", tmp1, 6 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV055_M16", tmp1, 6 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV056_M16", tmp1, 6 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV057_M16", tmp1, 6 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV058_M16", tmp1, 6 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV059_M16", tmp1, 6 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV060_M16", tmp1, 6 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV061_M16", tmp1, 7 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV062_M16", tmp1, 7 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV063_M16", tmp1, 7 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV064_M16", tmp1, 7 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV065_M16", tmp1, 7 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV066_M16", tmp1, 7 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV067_M16", tmp1, 7 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV068_M16", tmp1, 7 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV069_M16", tmp1, 7 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV070_M16", tmp1, 7 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV071_M16", tmp1, 8 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV072_M16", tmp1, 8 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV073_M16", tmp1, 8 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV074_M16", tmp1, 8 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV075_M16", tmp1, 8 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV076_M16", tmp1, 8 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV077_M16", tmp1, 8 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV078_M16", tmp1, 8 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV079_M16", tmp1, 8 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV080_M16", tmp1, 8 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV081_M16", tmp1, 9 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV082_M16", tmp1, 9 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV083_M16", tmp1, 9 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV084_M16", tmp1, 9 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV085_M16", tmp1, 9 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV086_M16", tmp1, 9 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV087_M16", tmp1, 9 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV088_M16", tmp1, 9 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV089_M16", tmp1, 9 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV090_M16", tmp1, 9 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV091_M16", tmp1, 10 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV092_M16", tmp1, 10 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV093_M16", tmp1, 10 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV094_M16", tmp1, 10 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV095_M16", tmp1, 10 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV096_M16", tmp1, 10 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV097_M16", tmp1, 10 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV098_M16", tmp1, 10 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV099_M16", tmp1, 10 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV100_M16", tmp1, 10 * 10 + 9 - 10)

End Function



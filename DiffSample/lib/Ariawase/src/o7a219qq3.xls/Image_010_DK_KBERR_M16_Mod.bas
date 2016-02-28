Attribute VB_Name = "Image_010_DK_KBERR_M16_Mod"

Option Explicit

Public Function DK_KBERR_M16_Process()

        Call PutImageInto_Common

' #### DK_KBV001_M16 ####

    Dim site As Long

    ' 0.画像情報インポート.DK_KBV001_M16
    Dim DK_KBERR_M16_Param As CParamPlane
    Dim DK_KBERR_M16_DevInfo As CDeviceConfigInfo
    Dim DK_KBERR_M16_Plane As CImgPlane
    Set DK_KBERR_M16_Param = TheParameterBank.Item("DK_KB_M16ImageTest_Acq1")
    Set DK_KBERR_M16_DevInfo = TheDeviceProfiler.ConfigInfo("DK_KB_M16ImageTest_Acq1")
        Call TheParameterBank.Delete("DK_KB_M16ImageTest_Acq1")
    Set DK_KBERR_M16_Plane = DK_KBERR_M16_Param.plane

    ' 2.Median.DK_KBV001_M16
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call MedianEx(DK_KBERR_M16_Plane, sPlane1, "Bayer2x4_ZONE3", 1, 5)

    ' 3.Median.DK_KBV001_M16
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane1)

    ' 4.Subtract(通常).DK_KBV001_M16
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call Subtract(DK_KBERR_M16_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane2, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane3, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane2)

    ' 5.LSB定義.DK_KBV001_M16
    Dim DK_KBERR_M16_LSB() As Double
     DK_KBERR_M16_LSB = DK_KBERR_M16_DevInfo.Lsb.AsDouble

    ' 6.SliceLevel生成_ShiroKobu_マージプロセス.DK_KBV001_M16
    Dim tmp_Slice1(nSite, (0.01 - 0.0001) / 0.0001) As Double
    Call MakeSliceLevel_ShiroKobu_1Step(tmp_Slice1, 0.0001, 0.01, 0.0001, DK_KBERR_M16_LSB, 15 / 30)

    ' 7.Count_ShiroKobu_マージプロセス.DK_KBV001_M16
    Dim tmp1(nSite, (0.01 - 0.0001) / 0.0001) As Double
    Call Count_ShiroKobu_marge(sPlane3, "Bayer2x4_ZONE2D", tmp_Slice1, tmp1, "Between", "Between")

    ' 9.PutTestResult_マージプロセス.DK_KBV001_M16
    Call ResultAdd_ShiroKobu("DK_KBV001_M16", tmp1, 1 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV002_M16", tmp1, 1 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV003_M16", tmp1, 1 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV004_M16", tmp1, 1 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV005_M16", tmp1, 1 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV006_M16", tmp1, 1 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV007_M16", tmp1, 1 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV008_M16", tmp1, 1 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV009_M16", tmp1, 1 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV010_M16", tmp1, 1 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV011_M16", tmp1, 2 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV012_M16", tmp1, 2 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV013_M16", tmp1, 2 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV014_M16", tmp1, 2 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV015_M16", tmp1, 2 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV016_M16", tmp1, 2 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV017_M16", tmp1, 2 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV018_M16", tmp1, 2 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV019_M16", tmp1, 2 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV020_M16", tmp1, 2 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV021_M16", tmp1, 3 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV022_M16", tmp1, 3 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV023_M16", tmp1, 3 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV024_M16", tmp1, 3 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV025_M16", tmp1, 3 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV026_M16", tmp1, 3 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV027_M16", tmp1, 3 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV028_M16", tmp1, 3 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV029_M16", tmp1, 3 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV030_M16", tmp1, 3 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV031_M16", tmp1, 4 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV032_M16", tmp1, 4 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV033_M16", tmp1, 4 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV034_M16", tmp1, 4 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV035_M16", tmp1, 4 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV036_M16", tmp1, 4 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV037_M16", tmp1, 4 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV038_M16", tmp1, 4 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV039_M16", tmp1, 4 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV040_M16", tmp1, 4 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV041_M16", tmp1, 5 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV042_M16", tmp1, 5 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV043_M16", tmp1, 5 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV044_M16", tmp1, 5 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV045_M16", tmp1, 5 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV046_M16", tmp1, 5 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV047_M16", tmp1, 5 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV048_M16", tmp1, 5 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV049_M16", tmp1, 5 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV050_M16", tmp1, 5 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV051_M16", tmp1, 6 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV052_M16", tmp1, 6 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV053_M16", tmp1, 6 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV054_M16", tmp1, 6 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV055_M16", tmp1, 6 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV056_M16", tmp1, 6 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV057_M16", tmp1, 6 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV058_M16", tmp1, 6 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV059_M16", tmp1, 6 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV060_M16", tmp1, 6 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV061_M16", tmp1, 7 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV062_M16", tmp1, 7 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV063_M16", tmp1, 7 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV064_M16", tmp1, 7 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV065_M16", tmp1, 7 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV066_M16", tmp1, 7 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV067_M16", tmp1, 7 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV068_M16", tmp1, 7 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV069_M16", tmp1, 7 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV070_M16", tmp1, 7 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV071_M16", tmp1, 8 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV072_M16", tmp1, 8 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV073_M16", tmp1, 8 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV074_M16", tmp1, 8 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV075_M16", tmp1, 8 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV076_M16", tmp1, 8 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV077_M16", tmp1, 8 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV078_M16", tmp1, 8 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV079_M16", tmp1, 8 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV080_M16", tmp1, 8 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV081_M16", tmp1, 9 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV082_M16", tmp1, 9 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV083_M16", tmp1, 9 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV084_M16", tmp1, 9 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV085_M16", tmp1, 9 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV086_M16", tmp1, 9 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV087_M16", tmp1, 9 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV088_M16", tmp1, 9 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV089_M16", tmp1, 9 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV090_M16", tmp1, 9 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV091_M16", tmp1, 10 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV092_M16", tmp1, 10 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV093_M16", tmp1, 10 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV094_M16", tmp1, 10 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV095_M16", tmp1, 10 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV096_M16", tmp1, 10 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV097_M16", tmp1, 10 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV098_M16", tmp1, 10 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV099_M16", tmp1, 10 * 10 + 8 - 10)

' #### DK_KBV100_M16 ####

    ' 0.画像情報インポート.DK_KBV100_M16

    ' 2.Median.DK_KBV100_M16

    ' 3.Median.DK_KBV100_M16

    ' 4.Subtract(通常).DK_KBV100_M16

    ' 5.LSB定義.DK_KBV100_M16

    ' 6.SliceLevel生成_ShiroKobu_マージプロセス.DK_KBV100_M16
    Dim tmp_Slice2(nSite, (0.03 - 0.01) / 0.0002) As Double
    Call MakeSliceLevel_ShiroKobu_1Step(tmp_Slice2, 0.01, 0.03, 0.0002, DK_KBERR_M16_LSB, 15 / 30)

    ' 7.Count_ShiroKobu_マージプロセス.DK_KBV100_M16
    Dim tmp2(nSite, (0.03 - 0.01) / 0.0002) As Double
    Call Count_ShiroKobu_marge(sPlane3, "Bayer2x4_ZONE2D", tmp_Slice2, tmp2, "Between", "Above")
        Call ReleasePlane(sPlane3)

    ' 9.PutTestResult_マージプロセス.DK_KBV100_M16
    Call ResultAdd_ShiroKobu("DK_KBV100_M16", tmp2, 1 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV102_M16", tmp2, 1 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV104_M16", tmp2, 1 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV106_M16", tmp2, 1 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV108_M16", tmp2, 1 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV110_M16", tmp2, 1 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV112_M16", tmp2, 1 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV114_M16", tmp2, 1 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV116_M16", tmp2, 1 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV118_M16", tmp2, 1 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV120_M16", tmp2, 2 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV122_M16", tmp2, 2 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV124_M16", tmp2, 2 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV126_M16", tmp2, 2 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV128_M16", tmp2, 2 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV130_M16", tmp2, 2 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV132_M16", tmp2, 2 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV134_M16", tmp2, 2 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV136_M16", tmp2, 2 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV138_M16", tmp2, 2 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV140_M16", tmp2, 3 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV142_M16", tmp2, 3 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV144_M16", tmp2, 3 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV146_M16", tmp2, 3 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV148_M16", tmp2, 3 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV150_M16", tmp2, 3 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV152_M16", tmp2, 3 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV154_M16", tmp2, 3 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV156_M16", tmp2, 3 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV158_M16", tmp2, 3 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV160_M16", tmp2, 4 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV162_M16", tmp2, 4 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV164_M16", tmp2, 4 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV166_M16", tmp2, 4 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV168_M16", tmp2, 4 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV170_M16", tmp2, 4 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV172_M16", tmp2, 4 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV174_M16", tmp2, 4 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV176_M16", tmp2, 4 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV178_M16", tmp2, 4 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV180_M16", tmp2, 5 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV182_M16", tmp2, 5 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV184_M16", tmp2, 5 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV186_M16", tmp2, 5 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV188_M16", tmp2, 5 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV190_M16", tmp2, 5 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV192_M16", tmp2, 5 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV194_M16", tmp2, 5 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV196_M16", tmp2, 5 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV198_M16", tmp2, 5 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV200_M16", tmp2, 6 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV202_M16", tmp2, 6 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV204_M16", tmp2, 6 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV206_M16", tmp2, 6 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV208_M16", tmp2, 6 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV210_M16", tmp2, 6 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV212_M16", tmp2, 6 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV214_M16", tmp2, 6 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV216_M16", tmp2, 6 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV218_M16", tmp2, 6 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV220_M16", tmp2, 7 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV222_M16", tmp2, 7 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV224_M16", tmp2, 7 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV226_M16", tmp2, 7 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV228_M16", tmp2, 7 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV230_M16", tmp2, 7 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV232_M16", tmp2, 7 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV234_M16", tmp2, 7 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV236_M16", tmp2, 7 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV238_M16", tmp2, 7 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV240_M16", tmp2, 8 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV242_M16", tmp2, 8 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV244_M16", tmp2, 8 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV246_M16", tmp2, 8 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV248_M16", tmp2, 8 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV250_M16", tmp2, 8 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV252_M16", tmp2, 8 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV254_M16", tmp2, 8 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV256_M16", tmp2, 8 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV258_M16", tmp2, 8 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV260_M16", tmp2, 9 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV262_M16", tmp2, 9 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV264_M16", tmp2, 9 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV266_M16", tmp2, 9 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV268_M16", tmp2, 9 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV270_M16", tmp2, 9 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV272_M16", tmp2, 9 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV274_M16", tmp2, 9 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV276_M16", tmp2, 9 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV278_M16", tmp2, 9 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV280_M16", tmp2, 10 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV282_M16", tmp2, 10 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV284_M16", tmp2, 10 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV286_M16", tmp2, 10 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV288_M16", tmp2, 10 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV290_M16", tmp2, 10 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV292_M16", tmp2, 10 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV294_M16", tmp2, 10 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV296_M16", tmp2, 10 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV298_M16", tmp2, 10 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV300_M16", tmp2, 11 * 10 + 0 - 10)

End Function



Attribute VB_Name = "Image_009_DK_KBERR_M10_Mod"

Option Explicit

Public Function DK_KBERR_M10_Process()

        Call PutImageInto_Common

' #### DK_KBV001_M10 ####

    Dim site As Long

    ' 0.画像情報インポート.DK_KBV001_M10
    Dim DK_KBERR_M10_Param As CParamPlane
    Dim DK_KBERR_M10_DevInfo As CDeviceConfigInfo
    Dim DK_KBERR_M10_Plane As CImgPlane
    Set DK_KBERR_M10_Param = TheParameterBank.Item("DK_KB_M10ImageTest_Acq1")
    Set DK_KBERR_M10_DevInfo = TheDeviceProfiler.ConfigInfo("DK_KB_M10ImageTest_Acq1")
        Call TheParameterBank.Delete("DK_KB_M10ImageTest_Acq1")
    Set DK_KBERR_M10_Plane = DK_KBERR_M10_Param.plane

    ' 2.Median.DK_KBV001_M10
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call MedianEx(DK_KBERR_M10_Plane, sPlane1, "Bayer2x4_ZONE3", 1, 5)

    ' 3.Median.DK_KBV001_M10
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane1)

    ' 4.Subtract(通常).DK_KBV001_M10
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call Subtract(DK_KBERR_M10_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane2, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane3, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane2)

    ' 5.LSB定義.DK_KBV001_M10
    Dim DK_KBERR_M10_LSB() As Double
     DK_KBERR_M10_LSB = DK_KBERR_M10_DevInfo.Lsb.AsDouble

    ' 6.SliceLevel生成_ShiroKobu_マージプロセス.DK_KBV001_M10
    Dim tmp_Slice1(nSite, (0.01 - 0.0001) / 0.0001) As Double
    Call MakeSliceLevel_ShiroKobu_1Step(tmp_Slice1, 0.0001, 0.01, 0.0001, DK_KBERR_M10_LSB, 15 / 30)

    ' 7.Count_ShiroKobu_マージプロセス.DK_KBV001_M10
    Dim tmp1(nSite, (0.01 - 0.0001) / 0.0001) As Double
    Call Count_ShiroKobu_marge(sPlane3, "Bayer2x4_ZONE2D", tmp_Slice1, tmp1, "Between", "Between")

    ' 9.PutTestResult_マージプロセス.DK_KBV001_M10
    Call ResultAdd_ShiroKobu("DK_KBV001_M10", tmp1, 1 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV002_M10", tmp1, 1 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV003_M10", tmp1, 1 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV004_M10", tmp1, 1 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV005_M10", tmp1, 1 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV006_M10", tmp1, 1 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV007_M10", tmp1, 1 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV008_M10", tmp1, 1 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV009_M10", tmp1, 1 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV010_M10", tmp1, 1 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV011_M10", tmp1, 2 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV012_M10", tmp1, 2 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV013_M10", tmp1, 2 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV014_M10", tmp1, 2 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV015_M10", tmp1, 2 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV016_M10", tmp1, 2 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV017_M10", tmp1, 2 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV018_M10", tmp1, 2 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV019_M10", tmp1, 2 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV020_M10", tmp1, 2 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV021_M10", tmp1, 3 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV022_M10", tmp1, 3 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV023_M10", tmp1, 3 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV024_M10", tmp1, 3 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV025_M10", tmp1, 3 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV026_M10", tmp1, 3 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV027_M10", tmp1, 3 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV028_M10", tmp1, 3 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV029_M10", tmp1, 3 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV030_M10", tmp1, 3 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV031_M10", tmp1, 4 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV032_M10", tmp1, 4 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV033_M10", tmp1, 4 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV034_M10", tmp1, 4 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV035_M10", tmp1, 4 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV036_M10", tmp1, 4 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV037_M10", tmp1, 4 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV038_M10", tmp1, 4 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV039_M10", tmp1, 4 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV040_M10", tmp1, 4 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV041_M10", tmp1, 5 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV042_M10", tmp1, 5 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV043_M10", tmp1, 5 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV044_M10", tmp1, 5 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV045_M10", tmp1, 5 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV046_M10", tmp1, 5 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV047_M10", tmp1, 5 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV048_M10", tmp1, 5 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV049_M10", tmp1, 5 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV050_M10", tmp1, 5 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV051_M10", tmp1, 6 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV052_M10", tmp1, 6 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV053_M10", tmp1, 6 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV054_M10", tmp1, 6 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV055_M10", tmp1, 6 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV056_M10", tmp1, 6 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV057_M10", tmp1, 6 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV058_M10", tmp1, 6 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV059_M10", tmp1, 6 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV060_M10", tmp1, 6 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV061_M10", tmp1, 7 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV062_M10", tmp1, 7 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV063_M10", tmp1, 7 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV064_M10", tmp1, 7 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV065_M10", tmp1, 7 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV066_M10", tmp1, 7 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV067_M10", tmp1, 7 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV068_M10", tmp1, 7 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV069_M10", tmp1, 7 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV070_M10", tmp1, 7 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV071_M10", tmp1, 8 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV072_M10", tmp1, 8 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV073_M10", tmp1, 8 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV074_M10", tmp1, 8 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV075_M10", tmp1, 8 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV076_M10", tmp1, 8 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV077_M10", tmp1, 8 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV078_M10", tmp1, 8 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV079_M10", tmp1, 8 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV080_M10", tmp1, 8 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV081_M10", tmp1, 9 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV082_M10", tmp1, 9 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV083_M10", tmp1, 9 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV084_M10", tmp1, 9 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV085_M10", tmp1, 9 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV086_M10", tmp1, 9 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV087_M10", tmp1, 9 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV088_M10", tmp1, 9 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV089_M10", tmp1, 9 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV090_M10", tmp1, 9 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV091_M10", tmp1, 10 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV092_M10", tmp1, 10 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV093_M10", tmp1, 10 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV094_M10", tmp1, 10 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV095_M10", tmp1, 10 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV096_M10", tmp1, 10 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV097_M10", tmp1, 10 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV098_M10", tmp1, 10 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV099_M10", tmp1, 10 * 10 + 8 - 10)

' #### DK_KBV100_M10 ####

    ' 0.画像情報インポート.DK_KBV100_M10

    ' 2.Median.DK_KBV100_M10

    ' 3.Median.DK_KBV100_M10

    ' 4.Subtract(通常).DK_KBV100_M10

    ' 5.LSB定義.DK_KBV100_M10

    ' 6.SliceLevel生成_ShiroKobu_マージプロセス.DK_KBV100_M10
    Dim tmp_Slice2(nSite, (0.03 - 0.01) / 0.0002) As Double
    Call MakeSliceLevel_ShiroKobu_1Step(tmp_Slice2, 0.01, 0.03, 0.0002, DK_KBERR_M10_LSB, 15 / 30)

    ' 7.Count_ShiroKobu_マージプロセス.DK_KBV100_M10
    Dim tmp2(nSite, (0.03 - 0.01) / 0.0002) As Double
    Call Count_ShiroKobu_marge(sPlane3, "Bayer2x4_ZONE2D", tmp_Slice2, tmp2, "Between", "Above")
        Call ReleasePlane(sPlane3)

    ' 9.PutTestResult_マージプロセス.DK_KBV100_M10
    Call ResultAdd_ShiroKobu("DK_KBV100_M10", tmp2, 1 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV102_M10", tmp2, 1 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV104_M10", tmp2, 1 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV106_M10", tmp2, 1 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV108_M10", tmp2, 1 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV110_M10", tmp2, 1 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV112_M10", tmp2, 1 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV114_M10", tmp2, 1 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV116_M10", tmp2, 1 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV118_M10", tmp2, 1 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV120_M10", tmp2, 2 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV122_M10", tmp2, 2 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV124_M10", tmp2, 2 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV126_M10", tmp2, 2 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV128_M10", tmp2, 2 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV130_M10", tmp2, 2 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV132_M10", tmp2, 2 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV134_M10", tmp2, 2 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV136_M10", tmp2, 2 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV138_M10", tmp2, 2 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV140_M10", tmp2, 3 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV142_M10", tmp2, 3 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV144_M10", tmp2, 3 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV146_M10", tmp2, 3 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV148_M10", tmp2, 3 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV150_M10", tmp2, 3 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV152_M10", tmp2, 3 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV154_M10", tmp2, 3 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV156_M10", tmp2, 3 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV158_M10", tmp2, 3 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV160_M10", tmp2, 4 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV162_M10", tmp2, 4 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV164_M10", tmp2, 4 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV166_M10", tmp2, 4 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV168_M10", tmp2, 4 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV170_M10", tmp2, 4 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV172_M10", tmp2, 4 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV174_M10", tmp2, 4 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV176_M10", tmp2, 4 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV178_M10", tmp2, 4 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV180_M10", tmp2, 5 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV182_M10", tmp2, 5 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV184_M10", tmp2, 5 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV186_M10", tmp2, 5 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV188_M10", tmp2, 5 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV190_M10", tmp2, 5 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV192_M10", tmp2, 5 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV194_M10", tmp2, 5 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV196_M10", tmp2, 5 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV198_M10", tmp2, 5 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV200_M10", tmp2, 6 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV202_M10", tmp2, 6 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV204_M10", tmp2, 6 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV206_M10", tmp2, 6 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV208_M10", tmp2, 6 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV210_M10", tmp2, 6 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV212_M10", tmp2, 6 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV214_M10", tmp2, 6 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV216_M10", tmp2, 6 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV218_M10", tmp2, 6 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV220_M10", tmp2, 7 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV222_M10", tmp2, 7 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV224_M10", tmp2, 7 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV226_M10", tmp2, 7 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV228_M10", tmp2, 7 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV230_M10", tmp2, 7 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV232_M10", tmp2, 7 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV234_M10", tmp2, 7 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV236_M10", tmp2, 7 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV238_M10", tmp2, 7 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV240_M10", tmp2, 8 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV242_M10", tmp2, 8 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV244_M10", tmp2, 8 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV246_M10", tmp2, 8 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV248_M10", tmp2, 8 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV250_M10", tmp2, 8 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV252_M10", tmp2, 8 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV254_M10", tmp2, 8 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV256_M10", tmp2, 8 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV258_M10", tmp2, 8 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV260_M10", tmp2, 9 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV262_M10", tmp2, 9 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV264_M10", tmp2, 9 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV266_M10", tmp2, 9 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV268_M10", tmp2, 9 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV270_M10", tmp2, 9 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV272_M10", tmp2, 9 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV274_M10", tmp2, 9 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV276_M10", tmp2, 9 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV278_M10", tmp2, 9 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV280_M10", tmp2, 10 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV282_M10", tmp2, 10 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV284_M10", tmp2, 10 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV286_M10", tmp2, 10 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV288_M10", tmp2, 10 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV290_M10", tmp2, 10 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV292_M10", tmp2, 10 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV294_M10", tmp2, 10 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV296_M10", tmp2, 10 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV298_M10", tmp2, 10 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV300_M10", tmp2, 11 * 10 + 0 - 10)

End Function



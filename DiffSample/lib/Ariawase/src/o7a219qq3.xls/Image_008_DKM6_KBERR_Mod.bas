Attribute VB_Name = "Image_008_DKM6_KBERR_Mod"

Option Explicit

Public Function DKM6_KBERR_Process()

        Call PutImageInto_Common

' #### DKM6_KBV010 ####

    Dim site As Long

    ' 0.画像情報インポート.DKM6_KBV010
    Dim DKM6_KBERR_Param As CParamPlane
    Dim DKM6_KBERR_DevInfo As CDeviceConfigInfo
    Dim DKM6_KBERR_Plane As CImgPlane
    Set DKM6_KBERR_Param = TheParameterBank.Item("DKM6_KBImageTest_Acq1")
    Set DKM6_KBERR_DevInfo = TheDeviceProfiler.ConfigInfo("DKM6_KBImageTest_Acq1")
        Call TheParameterBank.Delete("DKM6_KBImageTest_Acq1")
    Set DKM6_KBERR_Plane = DKM6_KBERR_Param.plane

    ' 2.Median.DKM6_KBV010
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call MedianEx(DKM6_KBERR_Plane, sPlane1, "Bayer2x4_ZONE3", 1, 5)

    ' 3.Median.DKM6_KBV010
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane1)

    ' 4.Subtract(通常).DKM6_KBV010
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call Subtract(DKM6_KBERR_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane2, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane3, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane2)

    ' 5.LSB定義.DKM6_KBV010
    Dim DKM6_KBERR_LSB() As Double
     DKM6_KBERR_LSB = DKM6_KBERR_DevInfo.Lsb.AsDouble

    ' 6.SliceLevel生成_ShiroKobu_マージプロセス.DKM6_KBV010
    Dim tmp_Slice1(nSite, (0.35 - 0.01) / 0.01) As Double
    Call MakeSliceLevel_ShiroKobu_1Step(tmp_Slice1, 0.01, 0.35, 0.01, DKM6_KBERR_LSB, 15 / 30)

    ' 7.Count_ShiroKobu_マージプロセス.DKM6_KBV010
    Dim tmp1(nSite, (0.35 - 0.01) / 0.01) As Double
    Call Count_ShiroKobu_marge(sPlane3, "Bayer2x4_ZONE2D", tmp_Slice1, tmp1, "Between", "Above")
        Call ReleasePlane(sPlane3)

    ' 8.PutTestResult_マージプロセス.DKM6_KBV010
    Call ResultAdd_ShiroKobu("DKM6_KBV010", tmp1, 1 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV020", tmp1, 1 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV030", tmp1, 1 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV040", tmp1, 1 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV050", tmp1, 1 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV060", tmp1, 1 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV070", tmp1, 1 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV080", tmp1, 1 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV090", tmp1, 1 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV100", tmp1, 1 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV110", tmp1, 2 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV120", tmp1, 2 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV130", tmp1, 2 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV140", tmp1, 2 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV150", tmp1, 2 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV160", tmp1, 2 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV170", tmp1, 2 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV180", tmp1, 2 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV190", tmp1, 2 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV200", tmp1, 2 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV210", tmp1, 3 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV220", tmp1, 3 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV230", tmp1, 3 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV240", tmp1, 3 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV250", tmp1, 3 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV260", tmp1, 3 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV270", tmp1, 3 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV280", tmp1, 3 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV290", tmp1, 3 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV300", tmp1, 3 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV310", tmp1, 4 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV320", tmp1, 4 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV330", tmp1, 4 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV340", tmp1, 4 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKM6_KBV350", tmp1, 4 * 10 + 4 - 10)

End Function



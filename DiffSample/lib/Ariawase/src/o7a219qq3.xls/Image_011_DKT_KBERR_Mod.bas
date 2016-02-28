Attribute VB_Name = "Image_011_DKT_KBERR_Mod"

Option Explicit

Public Function DKT_KBERR_Process()

        Call PutImageInto_Common

' #### DKT_KBV001 ####

    Dim site As Long

    ' 0.画像情報インポート.DKT_KBV001
    Dim DKT_KBERR_Param As CParamPlane
    Dim DKT_KBERR_DevInfo As CDeviceConfigInfo
    Dim DKT_KBERR_Plane As CImgPlane
    Set DKT_KBERR_Param = TheParameterBank.Item("DKT_KBImageTest_Acq1")
    Set DKT_KBERR_DevInfo = TheDeviceProfiler.ConfigInfo("DKT_KBImageTest_Acq1")
        Call TheParameterBank.Delete("DKT_KBImageTest_Acq1")
    Set DKT_KBERR_Plane = DKT_KBERR_Param.plane

    ' 2.Median.DKT_KBV001
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call MedianEx(DKT_KBERR_Plane, sPlane1, "Bayer2x4_ZONE3", 1, 5)

    ' 3.Median.DKT_KBV001
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane1)

    ' 4.Subtract(通常).DKT_KBV001
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call Subtract(DKT_KBERR_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane2, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane3, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane2)

    ' 5.LSB定義.DKT_KBV001
    Dim DKT_KBERR_LSB() As Double
     DKT_KBERR_LSB = DKT_KBERR_DevInfo.Lsb.AsDouble

    ' 6.SliceLevel生成_ShiroKobu_マージプロセス.DKT_KBV001
    Dim tmp_Slice1(nSite, (0.01 - 0.0001) / 0.0001) As Double
    Call MakeSliceLevel_ShiroKobu_1Step(tmp_Slice1, 0.0001, 0.01, 0.0001, DKT_KBERR_LSB)

    ' 7.Count_ShiroKobu_マージプロセス.DKT_KBV001
    Dim tmp1(nSite, (0.01 - 0.0001) / 0.0001) As Double
    Call Count_ShiroKobu_marge(sPlane3, "Bayer2x4_ZONE2D", tmp_Slice1, tmp1, "Between", "Above")
        Call ReleasePlane(sPlane3)

    ' 8.PutTestResult_マージプロセス.DKT_KBV001
    Call ResultAdd_ShiroKobu("DKT_KBV001", tmp1, 1 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV002", tmp1, 1 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV003", tmp1, 1 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV004", tmp1, 1 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV005", tmp1, 1 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV006", tmp1, 1 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV007", tmp1, 1 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV008", tmp1, 1 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV009", tmp1, 1 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV010", tmp1, 1 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV011", tmp1, 2 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV012", tmp1, 2 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV013", tmp1, 2 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV014", tmp1, 2 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV015", tmp1, 2 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV016", tmp1, 2 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV017", tmp1, 2 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV018", tmp1, 2 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV019", tmp1, 2 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV020", tmp1, 2 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV021", tmp1, 3 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV022", tmp1, 3 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV023", tmp1, 3 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV024", tmp1, 3 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV025", tmp1, 3 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV026", tmp1, 3 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV027", tmp1, 3 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV028", tmp1, 3 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV029", tmp1, 3 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV030", tmp1, 3 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV031", tmp1, 4 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV032", tmp1, 4 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV033", tmp1, 4 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV034", tmp1, 4 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV035", tmp1, 4 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV036", tmp1, 4 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV037", tmp1, 4 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV038", tmp1, 4 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV039", tmp1, 4 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV040", tmp1, 4 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV041", tmp1, 5 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV042", tmp1, 5 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV043", tmp1, 5 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV044", tmp1, 5 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV045", tmp1, 5 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV046", tmp1, 5 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV047", tmp1, 5 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV048", tmp1, 5 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV049", tmp1, 5 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV050", tmp1, 5 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV051", tmp1, 6 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV052", tmp1, 6 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV053", tmp1, 6 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV054", tmp1, 6 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV055", tmp1, 6 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV056", tmp1, 6 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV057", tmp1, 6 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV058", tmp1, 6 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV059", tmp1, 6 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV060", tmp1, 6 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV061", tmp1, 7 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV062", tmp1, 7 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV063", tmp1, 7 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV064", tmp1, 7 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV065", tmp1, 7 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV066", tmp1, 7 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV067", tmp1, 7 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV068", tmp1, 7 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV069", tmp1, 7 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV070", tmp1, 7 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV071", tmp1, 8 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV072", tmp1, 8 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV073", tmp1, 8 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV074", tmp1, 8 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV075", tmp1, 8 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV076", tmp1, 8 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV077", tmp1, 8 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV078", tmp1, 8 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV079", tmp1, 8 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV080", tmp1, 8 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV081", tmp1, 9 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV082", tmp1, 9 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV083", tmp1, 9 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV084", tmp1, 9 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV085", tmp1, 9 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV086", tmp1, 9 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV087", tmp1, 9 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV088", tmp1, 9 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV089", tmp1, 9 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV090", tmp1, 9 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV091", tmp1, 10 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV092", tmp1, 10 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV093", tmp1, 10 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV094", tmp1, 10 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV095", tmp1, 10 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV096", tmp1, 10 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV097", tmp1, 10 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV098", tmp1, 10 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV099", tmp1, 10 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DKT_KBV100", tmp1, 10 * 10 + 9 - 10)

End Function



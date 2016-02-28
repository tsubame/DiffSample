Attribute VB_Name = "Image_007_DK_KBERR_Mod"

Option Explicit

Public Function DK_KBERR_Process()

        Call PutImageInto_Common

' #### DK_KBV001 ####

    Dim site As Long

    ' 0.画像情報インポート.DK_KBV001
    Dim DK_KBERR_Param As CParamPlane
    Dim DK_KBERR_DevInfo As CDeviceConfigInfo
    Dim DK_KBERR_Plane As CImgPlane
    Set DK_KBERR_Param = TheParameterBank.Item("DK_KBImageTest_Acq1")
    Set DK_KBERR_DevInfo = TheDeviceProfiler.ConfigInfo("DK_KBImageTest_Acq1")
        Call TheParameterBank.Delete("DK_KBImageTest_Acq1")
    Set DK_KBERR_Plane = DK_KBERR_Param.plane

    ' 2.Median.DK_KBV001
    Dim sPlane1 As CImgPlane
    Call GetFreePlane(sPlane1, "Normal_Bayer2x4", idpDepthS16, , "sPlane1")
    Call MedianEx(DK_KBERR_Plane, sPlane1, "Bayer2x4_ZONE3", 1, 5)

    ' 3.Median.DK_KBV001
    Dim sPlane2 As CImgPlane
    Call GetFreePlane(sPlane2, "Normal_Bayer2x4", idpDepthS16, , "sPlane2")
    Call MedianEx(sPlane1, sPlane2, "Bayer2x4_ZONE3", 5, 1)
        Call ReleasePlane(sPlane1)

    ' 4.Subtract(通常).DK_KBV001
    Dim sPlane3 As CImgPlane
    Call GetFreePlane(sPlane3, "Normal_Bayer2x4", idpDepthS16, , "sPlane3")
    Call Subtract(DK_KBERR_Plane, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane2, "Bayer2x4_FULL", EEE_COLOR_ALL, sPlane3, "Bayer2x4_FULL", EEE_COLOR_ALL)
        Call ReleasePlane(sPlane2)

    ' 5.LSB定義.DK_KBV001
    Dim DK_KBERR_LSB() As Double
     DK_KBERR_LSB = DK_KBERR_DevInfo.Lsb.AsDouble

    ' 6.SliceLevel生成_ShiroKobu_マージプロセス.DK_KBV001
    Dim tmp_Slice1(nSite, (0.01 - 0.0001) / 0.0001) As Double
    Call MakeSliceLevel_ShiroKobu_1Step(tmp_Slice1, 0.0001, 0.01, 0.0001, DK_KBERR_LSB, 15 / 30)

    ' 7.Count_ShiroKobu_マージプロセス.DK_KBV001
    Dim tmp1(nSite, (0.01 - 0.0001) / 0.0001) As Double
    Call Count_ShiroKobu_marge(sPlane3, "Bayer2x4_ZONE2D", tmp_Slice1, tmp1, "Between", "Between")

    ' 9.PutTestResult_マージプロセス.DK_KBV001
    Call ResultAdd_ShiroKobu("DK_KBV001", tmp1, 1 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV002", tmp1, 1 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV003", tmp1, 1 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV004", tmp1, 1 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV005", tmp1, 1 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV006", tmp1, 1 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV007", tmp1, 1 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV008", tmp1, 1 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV009", tmp1, 1 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV010", tmp1, 1 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV011", tmp1, 2 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV012", tmp1, 2 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV013", tmp1, 2 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV014", tmp1, 2 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV015", tmp1, 2 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV016", tmp1, 2 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV017", tmp1, 2 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV018", tmp1, 2 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV019", tmp1, 2 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV020", tmp1, 2 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV021", tmp1, 3 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV022", tmp1, 3 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV023", tmp1, 3 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV024", tmp1, 3 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV025", tmp1, 3 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV026", tmp1, 3 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV027", tmp1, 3 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV028", tmp1, 3 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV029", tmp1, 3 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV030", tmp1, 3 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV031", tmp1, 4 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV032", tmp1, 4 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV033", tmp1, 4 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV034", tmp1, 4 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV035", tmp1, 4 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV036", tmp1, 4 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV037", tmp1, 4 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV038", tmp1, 4 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV039", tmp1, 4 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV040", tmp1, 4 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV041", tmp1, 5 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV042", tmp1, 5 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV043", tmp1, 5 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV044", tmp1, 5 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV045", tmp1, 5 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV046", tmp1, 5 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV047", tmp1, 5 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV048", tmp1, 5 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV049", tmp1, 5 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV050", tmp1, 5 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV051", tmp1, 6 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV052", tmp1, 6 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV053", tmp1, 6 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV054", tmp1, 6 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV055", tmp1, 6 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV056", tmp1, 6 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV057", tmp1, 6 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV058", tmp1, 6 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV059", tmp1, 6 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV060", tmp1, 6 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV061", tmp1, 7 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV062", tmp1, 7 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV063", tmp1, 7 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV064", tmp1, 7 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV065", tmp1, 7 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV066", tmp1, 7 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV067", tmp1, 7 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV068", tmp1, 7 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV069", tmp1, 7 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV070", tmp1, 7 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV071", tmp1, 8 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV072", tmp1, 8 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV073", tmp1, 8 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV074", tmp1, 8 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV075", tmp1, 8 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV076", tmp1, 8 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV077", tmp1, 8 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV078", tmp1, 8 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV079", tmp1, 8 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV080", tmp1, 8 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV081", tmp1, 9 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV082", tmp1, 9 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV083", tmp1, 9 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV084", tmp1, 9 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV085", tmp1, 9 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV086", tmp1, 9 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV087", tmp1, 9 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV088", tmp1, 9 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV089", tmp1, 9 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV090", tmp1, 9 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV091", tmp1, 10 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV092", tmp1, 10 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV093", tmp1, 10 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV094", tmp1, 10 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV095", tmp1, 10 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV096", tmp1, 10 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV097", tmp1, 10 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV098", tmp1, 10 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV099", tmp1, 10 * 10 + 8 - 10)

' #### DK_KBV100 ####

    ' 0.画像情報インポート.DK_KBV100

    ' 2.Median.DK_KBV100

    ' 3.Median.DK_KBV100

    ' 4.Subtract(通常).DK_KBV100

    ' 5.LSB定義.DK_KBV100

    ' 6.SliceLevel生成_ShiroKobu_マージプロセス.DK_KBV100
    Dim tmp_Slice2(nSite, (0.03 - 0.01) / 0.0002) As Double
    Call MakeSliceLevel_ShiroKobu_1Step(tmp_Slice2, 0.01, 0.03, 0.0002, DK_KBERR_LSB, 15 / 30)

    ' 7.Count_ShiroKobu_マージプロセス.DK_KBV100
    Dim tmp2(nSite, (0.03 - 0.01) / 0.0002) As Double
    Call Count_ShiroKobu_marge(sPlane3, "Bayer2x4_ZONE2D", tmp_Slice2, tmp2, "Between", "Above")
        Call ReleasePlane(sPlane3)

    ' 9.PutTestResult_マージプロセス.DK_KBV100
    Call ResultAdd_ShiroKobu("DK_KBV100", tmp2, 1 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV102", tmp2, 1 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV104", tmp2, 1 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV106", tmp2, 1 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV108", tmp2, 1 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV110", tmp2, 1 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV112", tmp2, 1 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV114", tmp2, 1 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV116", tmp2, 1 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV118", tmp2, 1 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV120", tmp2, 2 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV122", tmp2, 2 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV124", tmp2, 2 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV126", tmp2, 2 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV128", tmp2, 2 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV130", tmp2, 2 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV132", tmp2, 2 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV134", tmp2, 2 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV136", tmp2, 2 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV138", tmp2, 2 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV140", tmp2, 3 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV142", tmp2, 3 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV144", tmp2, 3 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV146", tmp2, 3 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV148", tmp2, 3 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV150", tmp2, 3 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV152", tmp2, 3 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV154", tmp2, 3 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV156", tmp2, 3 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV158", tmp2, 3 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV160", tmp2, 4 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV162", tmp2, 4 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV164", tmp2, 4 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV166", tmp2, 4 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV168", tmp2, 4 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV170", tmp2, 4 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV172", tmp2, 4 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV174", tmp2, 4 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV176", tmp2, 4 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV178", tmp2, 4 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV180", tmp2, 5 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV182", tmp2, 5 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV184", tmp2, 5 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV186", tmp2, 5 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV188", tmp2, 5 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV190", tmp2, 5 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV192", tmp2, 5 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV194", tmp2, 5 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV196", tmp2, 5 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV198", tmp2, 5 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV200", tmp2, 6 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV202", tmp2, 6 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV204", tmp2, 6 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV206", tmp2, 6 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV208", tmp2, 6 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV210", tmp2, 6 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV212", tmp2, 6 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV214", tmp2, 6 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV216", tmp2, 6 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV218", tmp2, 6 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV220", tmp2, 7 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV222", tmp2, 7 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV224", tmp2, 7 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV226", tmp2, 7 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV228", tmp2, 7 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV230", tmp2, 7 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV232", tmp2, 7 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV234", tmp2, 7 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV236", tmp2, 7 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV238", tmp2, 7 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV240", tmp2, 8 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV242", tmp2, 8 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV244", tmp2, 8 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV246", tmp2, 8 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV248", tmp2, 8 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV250", tmp2, 8 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV252", tmp2, 8 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV254", tmp2, 8 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV256", tmp2, 8 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV258", tmp2, 8 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV260", tmp2, 9 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV262", tmp2, 9 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV264", tmp2, 9 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV266", tmp2, 9 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV268", tmp2, 9 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV270", tmp2, 9 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV272", tmp2, 9 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV274", tmp2, 9 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV276", tmp2, 9 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV278", tmp2, 9 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV280", tmp2, 10 * 10 + 0 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV282", tmp2, 10 * 10 + 1 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV284", tmp2, 10 * 10 + 2 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV286", tmp2, 10 * 10 + 3 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV288", tmp2, 10 * 10 + 4 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV290", tmp2, 10 * 10 + 5 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV292", tmp2, 10 * 10 + 6 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV294", tmp2, 10 * 10 + 7 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV296", tmp2, 10 * 10 + 8 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV298", tmp2, 10 * 10 + 9 - 10)
    Call ResultAdd_ShiroKobu("DK_KBV300", tmp2, 11 * 10 + 0 - 10)

End Function



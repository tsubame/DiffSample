Attribute VB_Name = "Image_004_DKT_RVLERR_Mod"
Option Explicit

Public Function DKT_RVLERR_Process()

        Call PutImageInto_Common

' #### DKT_RVLSGM ####

        Dim site As Long

        Dim DKT_RVLERR_Param As CParamPlane
        Dim DKT_RVLERR_DevInfo As CDeviceConfigInfo
        Dim DKT_RVLERR_Plane As CImgPlane
        Set DKT_RVLERR_Param = TheParameterBank.Item("DKT_RVLImageTest_Acq1")
        Set DKT_RVLERR_DevInfo = DKT_RVLERR_Param.DeviceConfigInfo
        Set DKT_RVLERR_Plane = DKT_RVLERR_Param.plane

        Dim DKT_RVLERR_LSB() As Double
        DKT_RVLERR_LSB = DKT_RVLERR_DevInfo.Lsb.AsDouble

        Dim MedianPlane1 As CImgPlane
        Dim subtractPlane1 As CImgPlane
       
        Call GetFreePlane(MedianPlane1, DKT_RVLERR_Plane.planeGroup, idpDepthS16, True, "MedianPlane1")
        Call GetFreePlane(subtractPlane1, DKT_RVLERR_Plane.planeGroup, idpDepthS16, True, "subtractPlane1")
        
        Call MedianEx(DKT_RVLERR_Plane, MedianPlane1, "Bayer2x4_ZONE2D", 1, 5)

        Dim tmpDKMask(nSite) As Double
'        Dim tmpDKMask As CImgColorAllResult
        Dim maskPlane As CImgPlane
        Call GetFreePlane(maskPlane, DKT_RVLERR_Plane.planeGroup, idpDepthS16, True, "maskPlane")
        Call Copy(DKT_RVLERR_Plane, "Bayer2x4_FULL", EEE_COLOR_FLAT, maskPlane, "Bayer2x4_FULL", EEE_COLOR_FLAT)
        Call Subtract(DKT_RVLERR_Plane, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, MedianPlane1, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, subtractPlane1, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT)
        Call Count(subtractPlane1, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, idpCountOutside, -10, 10, idpLimitInclude, tmpDKMask, "FLG_DKDEF")

        With maskPlane
            Call .SetPMD("Bayer2x4_ZONE2D")
            Call .WritePixel(64, idpColorFlat, , "FLG_DKDEF")
        End With
        Call maskPlane.GetSharedFlagPlane("FLG_DKDEF").RemoveFlagBit("FLG_DKDEF")

        Dim Accplane1 As CImgPlane
'        Dim SubColumnsplane1 As CImgPlane
        Call GetFreePlane(Accplane1, DKT_RVLERR_Plane.planeGroup, idpDepthF32, True, "Accplane1")
'        Call GetFreePlane(SubColumnsplane1, DKT_RVLERR_Plane.planeGroup, idpDepthF32, True, "SubColumnsplane1")
        Call AccumulateColumn(maskPlane, "Bayer2x4_ZONE2D", EEE_COLOR_FLAT, Accplane1, "Bayer2x4_ZONE2D_VLINE", EEE_COLOR_FLAT, idpStdDeviation, 1)

'        'BIBUN -> SubColumnsplane1
'        Call SubColumns(Accplane1, "Bayer2x4_ZONE2D_VLINE", EEE_COLOR_FLAT, SubColumnsplane1, "Bayer2x4_ZONE2D_VLINE", EEE_COLOR_FLAT, 2)
'
'        'BIBUN ARI
'        Dim tmpAbsmax1(nSite) As Double
'        Dim tmpDKVLN1Result(nSite) As Double
'        Call AbsMax(SubColumnsplane1, "Bayer2x4_ZONE2D_VLINE", EEE_COLOR_FLAT, tmpAbsmax1)
'        For site = 0 To nSite
'            If TheExec.Sites.site(site).Active Then
'                tmpDKVLN1Result(site) = (Abs(tmpAbsmax1(site)) * DKT_RVLERR_LSB(site))
'            End If
'        Next site
'
'        Call ResultAdd("DK_VLN", tmpDKVLN1Result)

        'BIBUN NASHI
        Dim tmpAbsmax2(nSite) As Double
        Dim tmpDKVLN2Result(nSite) As Double
        Call AbsMax(Accplane1, "Bayer2x4_ZONE2D_VLINE", EEE_COLOR_FLAT, tmpAbsmax2)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                tmpDKVLN2Result(site) = tmpAbsmax2(site) * DKT_RVLERR_LSB(site)
            End If
        Next site

        Call ResultAdd("DKT_RVLSGM", tmpDKVLN2Result)

        Call ReleasePlane(MedianPlane1)
        Call ReleasePlane(subtractPlane1)
        Call ReleasePlane(maskPlane)
'        Call ReleasePlane(SubColumnsplane1)
        Call ReleasePlane(Accplane1)

End Function


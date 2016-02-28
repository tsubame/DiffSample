Attribute VB_Name = "Image_MasterFunctions"
Option Explicit

Private Type pmdInfo
    Left As Long
    Top As Long
    width As Long
    height As Long
    Right As Long
    Bottom As Long
End Type

Public Function GetMin(ByRef OutVar() As Double, ParamArray dataArr() As Variant)
    
    Dim site As Long
    Dim i As Long
    Dim tmp(nSite) As Double
    
    For site = 0 To nSite
    
        tmp(site) = dataArr(0)(site)
        
        For i = 1 To UBound(dataArr)
            If tmp(site) > dataArr(i)(site) Then
                tmp(site) = dataArr(i)(site)
            End If
        Next i
        
        OutVar(site) = tmp(site)
        
    Next site

End Function

Public Function GetMax(ByRef OutVar() As Double, ParamArray dataArr() As Variant)

    Dim site As Long
    Dim i As Long
    Dim tmp(nSite) As Double
        
    For site = 0 To nSite
    
        tmp(site) = dataArr(0)(site)
        
        For i = 1 To UBound(dataArr)
            If tmp(site) < dataArr(i)(site) Then
                tmp(site) = dataArr(i)(site)
            End If
        Next i
        
        OutVar(site) = tmp(site)
        
    Next site
    
End Function

Public Function GetAbsMax(ByRef OutVar() As Double, ParamArray dataArr() As Variant)

    Dim site As Long
    Dim i As Long
    Dim tmp(nSite) As Double
        
    For site = 0 To nSite
    
        tmp(site) = dataArr(0)(site)
        For i = 1 To UBound(dataArr)
            If Abs(tmp(site)) < Abs(dataArr(i)(site)) Then
                tmp(site) = dataArr(i)(site)
            End If
        Next i
        
        OutVar(site) = tmp(site)
        
    Next site

End Function
Public Function GetAbs(ByRef OutVar() As Double, ByRef inVar() As Double)

    Dim site As Long
        
    For site = 0 To nSite
        OutVar(site) = Abs(inVar(site))
    Next site

End Function

Public Function GetSum(ByRef OutVar() As Double, ParamArray dataArr() As Variant)

    Dim site As Long
    Dim i As Long
    Dim tmp(nSite) As Double
        
    For site = 0 To nSite
    
        tmp(site) = dataArr(0)(site)
        For i = 1 To UBound(dataArr)
            tmp(site) = tmp(site) + dataArr(i)(site)
        Next i
        
        OutVar(site) = tmp(site)
        
    Next site

End Function

Public Function GetAverage(ByRef OutVar() As Double, ParamArray dataArr() As Variant)

    Dim site As Long
    Dim i As Long
    Dim tmp(nSite) As Double
        
    For site = 0 To nSite
    
        tmp(site) = dataArr(0)(site)
        For i = 1 To UBound(dataArr)
            tmp(site) = tmp(site) + dataArr(i)(site)
        Next i
        
        OutVar(site) = tmp(site) / (UBound(dataArr) + 1)
        
    Next site

End Function

Public Function GetSubtract(ByRef OutVar() As Double, inVar1() As Double, inVar2() As Double)

    Dim site As Long
    Dim i As Long
    Dim tmp(nSite) As Double
        
    For site = 0 To nSite
        OutVar(site) = inVar1(site) - inVar2(site)
    Next site

End Function

Public Function WritePixelAddr(ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef pDataArray() As T_PIXINFO, Optional ByVal pSite As Long = -1)

    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.WritePixelAddr(pDataArray, pSite)

End Function

Public Function GetMin_Color(ByRef OutVar() As Double, ByVal inVar As CImgColorAllResult, ParamArray dataArr() As Variant)

    Dim site As Long
    Dim i As Long
    Dim tmp(nSite) As Double

    ' ==== Add(130906 130343) ====
    Dim mf_colorList() As String
    mf_colorList = GetColorList(inVar.ColorList, dataArr)
    ' replace dataArr => mf_colorList
    ' ==== Add(130906 130343) ====


    For site = 0 To nSite
        tmp(site) = inVar.color(mf_colorList(0)).SiteValue(site)
        For i = 1 To UBound(mf_colorList)
            If tmp(site) > inVar.color(mf_colorList(i)).SiteValue(site) Then
                tmp(site) = inVar.color(mf_colorList(i)).SiteValue(site)
            End If
        Next i

        OutVar(site) = tmp(site)

    Next site

End Function

Public Function GetMax_Color(ByRef OutVar() As Double, ByVal inVar As CImgColorAllResult, ParamArray dataArr() As Variant)

    Dim site As Long
    Dim i As Long
    Dim tmp(nSite) As Double

    ' ==== Add(130906 130343) ====
    Dim mf_colorList() As String
    mf_colorList = GetColorList(inVar.ColorList, dataArr)
    ' replace dataArr => mf_colorList
    ' ==== Add(130906 130343) ====


    For site = 0 To nSite

        tmp(site) = inVar.color(mf_colorList(0)).SiteValue(site)
        For i = 1 To UBound(mf_colorList)
            If tmp(site) < inVar.color(mf_colorList(i)).SiteValue(site) Then
                tmp(site) = inVar.color(mf_colorList(i)).SiteValue(site)
            End If
        Next i

        OutVar(site) = tmp(site)

    Next site

End Function

Public Function GetAbsMax_Color(ByRef OutVar() As Double, ByVal inVar As CImgColorAllResult, ParamArray dataArr() As Variant)

    Dim site As Long
    Dim i As Long
    Dim tmp(nSite) As Double

    ' ==== Add(130906 130343) ====
    Dim mf_colorList() As String
    mf_colorList = GetColorList(inVar.ColorList, dataArr)
    ' replace dataArr => mf_colorList
    ' ==== Add(130906 130343) ====

    For site = 0 To nSite

        tmp(site) = inVar.color(mf_colorList(0)).SiteValue(site)
        For i = 1 To UBound(mf_colorList)
            If Abs(tmp(site)) < Abs(inVar.color(mf_colorList(i)).SiteValue(site)) Then
                tmp(site) = inVar.color(mf_colorList(i)).SiteValue(site)
            End If
        Next i

        OutVar(site) = tmp(site)

    Next site

End Function

Public Function GetSum_Color(ByRef OutVar() As Double, ByVal inVar As CImgColorAllResult, ParamArray dataArr() As Variant)

    Dim site As Long
    Dim i As Long
    Dim tmp(nSite) As Double
    ' ==== Add(130906 130343) ====
    Dim mf_colorList() As String
    mf_colorList = GetColorList(inVar.ColorList, dataArr)
    ' replace dataArr => mf_colorList
    ' ==== Add(130906 130343) ====

    For site = 0 To nSite

        tmp(site) = inVar.color(mf_colorList(0)).SiteValue(site)
        For i = 1 To UBound(mf_colorList)
            tmp(site) = tmp(site) + inVar.color(mf_colorList(i)).SiteValue(site)
        Next i

        OutVar(site) = tmp(site)

    Next site

End Function

Public Function GetAverage_Color(ByRef OutVar() As Double, ByVal inVar As CImgColorAllResult, ParamArray dataArr() As Variant)

    Dim site As Long
    Dim i As Long
    Dim tmp(nSite) As Double
    ' ==== Add(130906 130343) ====
    Dim mf_colorList() As String
    mf_colorList = GetColorList(inVar.ColorList, dataArr)
    ' ==== Add(130906 130343) ====

    For site = 0 To nSite

        tmp(site) = inVar.color(mf_colorList(0)).SiteValue(site)
        For i = 1 To UBound(mf_colorList)
            tmp(site) = tmp(site) + inVar.color(mf_colorList(i)).SiteValue(site)
        Next i

        OutVar(site) = tmp(site) / (UBound(mf_colorList) + 1)

    Next site

End Function

Public Function GetMin_CImgColor(ByRef OutVar As CImgColorAllResult, ParamArray dataArr() As Variant)

    Dim i As Long, j As Long
    Dim site As Long
    Dim tmp0 As Double
    Dim tmp As CImgColorAllResult
    Dim tmp1 As CImgColorAllResult
    Dim tmp2 As CImgColorAllResult
    Dim ColorMap() As String
    
    Set tmp = dataArr(0)
    Set tmp1 = tmp.Clone
    ColorMap = tmp1.ColorList
    
    For i = 1 To UBound(dataArr)
        Set tmp2 = dataArr(i)
        For j = 0 To UBound(ColorMap)
            For site = 0 To nSite
                If tmp1.color(ColorMap(j)).SiteValue(site) > tmp2.color(ColorMap(j)).SiteValue(site) Then
                    tmp0 = tmp2.color(ColorMap(j)).SiteValue(site)
                    Call tmp1.SetData(ColorMap(j), site, tmp0)
                End If
            Next site
        Next j
    Next i
    
    Set OutVar = tmp1
    
End Function

Public Function GetMax_CImgColor(ByRef OutVar As CImgColorAllResult, ParamArray dataArr() As Variant)

    Dim i As Long, j As Long
    Dim site As Long
    Dim tmp0 As Double
    Dim tmp As CImgColorAllResult
    Dim tmp1 As CImgColorAllResult
    Dim tmp2 As CImgColorAllResult
    Dim ColorMap() As String
    
    Set tmp = dataArr(0)
    Set tmp1 = tmp.Clone
    ColorMap = tmp1.ColorList
    
    For i = 1 To UBound(dataArr)
        Set tmp2 = dataArr(i)
        For j = 0 To UBound(ColorMap)
            For site = 0 To nSite
                If tmp1.color(ColorMap(j)).SiteValue(site) < tmp2.color(ColorMap(j)).SiteValue(site) Then
                    tmp0 = tmp2.color(ColorMap(j)).SiteValue(site)
                    Call tmp1.SetData(ColorMap(j), site, tmp0)
                End If
            Next site
        Next j
    Next i
    
    Set OutVar = tmp1
    
End Function


Public Function GetAbsMax_CImgColor(ByRef OutVar As CImgColorAllResult, ParamArray dataArr() As Variant)

    Dim i As Long, j As Long
    Dim site As Long
    Dim tmp0 As Double
    Dim tmp As CImgColorAllResult
    Dim tmp1 As CImgColorAllResult
    Dim tmp2 As CImgColorAllResult
    Dim ColorMap() As String
    
    Set tmp = dataArr(0)
    Set tmp1 = tmp.Clone
    ColorMap = tmp1.ColorList
    
    For i = 1 To UBound(dataArr)
        Set tmp2 = dataArr(i)
        For j = 0 To UBound(ColorMap)
            For site = 0 To nSite
                If Abs(tmp1.color(ColorMap(j)).SiteValue(site)) < Abs(tmp2.color(ColorMap(j)).SiteValue(site)) Then
                    tmp0 = tmp2.color(ColorMap(j)).SiteValue(site)
                    Call tmp1.SetData(ColorMap(j), site, tmp0)
                End If
            Next site
        Next j
    Next i
    
    Set OutVar = tmp1
    
End Function

Public Function GetSum_CImgColor(ByRef OutVar As CImgColorAllResult, ParamArray dataArr() As Variant)

    Dim i As Long, j As Long
    Dim site As Long
    Dim tmp0 As Double
    Dim tmp As CImgColorAllResult
    Dim tmp1 As CImgColorAllResult
    Dim tmp2 As CImgColorAllResult
    Dim ColorMap() As String
    
    Set tmp = dataArr(0)
    Set tmp1 = tmp.Clone
    ColorMap = tmp1.ColorList
    
    For i = 1 To UBound(dataArr)
        Set tmp2 = dataArr(i)
        For j = 0 To UBound(ColorMap)
            For site = 0 To nSite
                tmp0 = tmp1.color(ColorMap(j)).SiteValue(site) + tmp2.color(ColorMap(j)).SiteValue(site)
                Call tmp1.SetData(ColorMap(j), site, tmp0)
            Next site
        Next j
    Next i
    
    Set OutVar = tmp1
    
End Function

Public Function GetAverage_CImgColor(ByRef OutVar As CImgColorAllResult, ParamArray dataArr() As Variant)

    Dim i As Long, j As Long
    Dim site As Long
    Dim tmp As CImgColorAllResult
    Dim tmp1 As CImgColorAllResult
    Dim tmp2 As CImgColorAllResult
    Dim tmp3 As Double
    Dim tmp4 As Double
    Dim ColorMap() As String
    
    Set tmp = dataArr(0)
    Set tmp1 = tmp.Clone
    ColorMap = tmp1.ColorList
    
    For i = 1 To UBound(dataArr)
        Set tmp2 = dataArr(i)
        For j = 0 To UBound(ColorMap)
            For site = 0 To nSite
                tmp3 = tmp1.color(ColorMap(j)).SiteValue(site) + tmp2.color(ColorMap(j)).SiteValue(site)
                Call tmp1.SetData(ColorMap(j), site, tmp3)
            Next site
        Next j
    Next i
    
    For j = 0 To UBound(ColorMap)
        For site = 0 To nSite
            tmp4 = tmp1.color(ColorMap(j)).SiteValue(site) / (UBound(dataArr) + 1)
            Call tmp1.SetData(ColorMap(j), site, tmp4)
        Next site
    Next j
    
    Set OutVar = tmp1
    
End Function

Public Function GetSubtract_CImgColor(ByRef OutVar As CImgColorAllResult, ParamArray dataArr() As Variant)

    Dim i As Long, j As Long
    Dim site As Long
    Dim tmp0 As Double
    Dim tmp As CImgColorAllResult
    Dim tmp1 As CImgColorAllResult
    Dim tmp2 As CImgColorAllResult
    Dim ColorMap() As String
    
    Set tmp = dataArr(0)
    Set tmp1 = tmp.Clone
    Set tmp2 = dataArr(1)
    ColorMap = tmp1.ColorList

    For j = 0 To UBound(ColorMap)
        For site = 0 To nSite
            tmp0 = tmp1.color(ColorMap(j)).SiteValue(site) - tmp2.color(ColorMap(j)).SiteValue(site)
            Call tmp1.SetData(ColorMap(j), site, tmp0)
        Next site
    Next j
    
    Set OutVar = tmp1
    
End Function

Public Function GetMaxArr(ByRef inVar() As Double) As Double

    Dim i As Long
    Dim tmp As Double
    
    tmp = inVar(0)
    If UBound(inVar) > 0 Then
        For i = 1 To UBound(inVar)
            If tmp < inVar(i) Then
                tmp = inVar(i)
            End If
        Next i
    End If

    GetMaxArr = tmp

End Function

Public Function Div(ByVal inVar1 As Double, ByVal inVar2 As Double, ByVal ErrorCode As Double) As Double

    If inVar2 = 0 Then
        Div = ErrorCode
    Else
        Div = inVar1 / inVar2
    End If

End Function

Public Sub Clamp(ByRef srcPlane As CImgPlane, ByRef dstPlane As CImgPlane, ByVal clampZone As Variant)
'内容:
'   対象画像をOPBクランプする。
'
'[srcPlane]    IN   CImgPlane型:    元プレーン
'[dstPlane]    OUT  CImgPlane型:    結果格納プレーン
'[clampZone]   IN   String型:       OPBクランプのゾーン指定

    'エラー処理
    If srcPlane.planeGroup <> dstPlane.planeGroup Then
        MsgBox "SelectError! @Clamp"
    End If


    Dim Bclamp_ColorAll As CImgColorAllResult
    Dim Bclamp(nSite) As Double
    
    '========== ZONEOPB MEAN VALUE ========================
    Call Average(srcPlane, clampZone, EEE_COLOR_ALL, Bclamp)
    Call SubtractConst(srcPlane, srcPlane.BasePMD.Name, EEE_COLOR_ALL, Bclamp, dstPlane, dstPlane.BasePMD.Name, EEE_COLOR_ALL)
 
End Sub

Public Sub MedianEx(ByRef srcPlane As CImgPlane, ByRef dstPlane As CImgPlane, ByVal Zone As Variant, ByVal medianTap_H As Double, ByVal medianTap_V As Double)
    
    'エラー処理
    If medianTap_H <> 1 And medianTap_V <> 1 Then
        MsgBox "SelectError! @MedianEx"
    End If

    Dim Zone_H As Double
    Dim Zone_V As Double
    Dim tmp_H As Double
    Dim tmp_V As Double
    Dim CM_H As Double
    Dim CM_V As Double
    
    Zone_H = TheIDP.PMD(Zone).width
    Zone_V = TheIDP.PMD(Zone).height
    CM_H = srcPlane.planeMap.width
    CM_V = srcPlane.planeMap.height
    tmp_H = 2 / 3 * Zone_H / CM_H + 1
    tmp_V = 2 / 3 * Zone_V / CM_V + 1
    
    If tmp_H <= medianTap_H Then
        If Int(tmp_H) Mod 2 = 1 Then
            medianTap_H = Int(tmp_H)
        Else
            medianTap_H = Int(tmp_H) - 1
        End If
    End If
    If tmp_V <= medianTap_V Then
        If Int(tmp_V) Mod 2 = 1 Then
            medianTap_V = Int(tmp_V)
        Else
            medianTap_V = Int(tmp_V) - 1
        End If
    End If
    
    Call Median(srcPlane, Zone, EEE_COLOR_ALL, dstPlane, Zone, EEE_COLOR_ALL, medianTap_H, medianTap_V)
    Call Extention(dstPlane, Zone, dstPlane, -Int(medianTap_H / 2), -Int(medianTap_H / 2), -Int(medianTap_V / 2), -Int(medianTap_V / 2), EEE_COLOR_ALL)

End Sub

Public Sub EXMedian(ByRef srcPlane As CImgPlane, ByRef dstPlane As CImgPlane, _
                            ByVal Zone As String, ByVal medianTap_H As Long, ByVal medianTap_V As Long)
                            
    'エラー処理
    If medianTap_H <> 1 And medianTap_V <> 1 Then
        MsgBox "SelectError! @ExMedian"
    End If

    Dim pColor As String
    pColor = "EEE_COLOR_ALL"
    
    Dim exPlane As CImgPlane, tmpPlane As CImgPlane
    Call GetFreePlane(exPlane, srcPlane.planeGroup, srcPlane.BitDepth, True, "exPlane")
    Call GetFreePlane(tmpPlane, srcPlane.planeGroup, srcPlane.BitDepth, True, "tmpPlane")

    Dim CM_H As Long
    Dim CM_V As Long
    CM_H = srcPlane.planeMap.width
    CM_V = srcPlane.planeMap.height
    
    Dim ExBit_H As Long
    Dim ExBit_V As Long
    If pColor = "EEE_COLOR_FLAT" Then
        ExBit_H = Int(medianTap_H / 2)
        ExBit_V = Int(medianTap_V / 2)
    Else
        ExBit_H = Int(medianTap_H / 2) * CM_H
        ExBit_V = Int(medianTap_V / 2) * CM_V
    End If
    
    If TheIDP.PMD(Zone).Left < ExBit_H Or TheIDP.PMD(Zone).Top < ExBit_V Then
        Call srcPlane.SetPMD(Zone)
        With TheIDP.PMD(Zone)
            Call exPlane.SetCustomPMD((.Left + ExBit_H), (.Top + ExBit_V), .width, .height)
        End With
        Call exPlane.CopyPlane(srcPlane, pColor)
    Else
        Call srcPlane.SetPMD(Zone)
        With TheIDP.PMD(Zone)
            Call exPlane.SetCustomPMD(.Left, .Top, .width, .height)
        End With
        Call exPlane.CopyPlane(srcPlane, pColor)
    End If
        
    Dim ExPmdInfo As pmdInfo
    Call GetCurrentPmdInfo(exPlane, ExPmdInfo)
    
    Dim i As Long
    Dim ExLoop_V As Long
    Dim ExLoop_H As Long
    ExLoop_V = ExBit_V / CM_V
    ExLoop_H = ExBit_H / CM_H
    
    With ExPmdInfo
        '-------- EXTENTION ---------
        '----- Base Copy -----
        Call exPlane.SetCustomPMD(.Left, .Top, .width, .height)
        Call tmpPlane.SetCustomPMD(.Left, .Top, .width, .height)
        Call tmpPlane.CopyPlane(exPlane, pColor)

        If ExBit_H <> 0 Then
            '----- Left Copy -----
            Call exPlane.SetCustomPMD(.Left, .Top, CM_H, .height)
            For i = 1 To ExLoop_H
                Call tmpPlane.SetCustomPMD((.Left - CM_H * i), .Top, CM_H, .height)
                Call tmpPlane.CopyPlane(exPlane, pColor)
            Next i
            '----- Right Copy -----
            Call exPlane.SetCustomPMD((.Right - CM_H + 1), .Top, CM_H, .height)
            For i = 1 To ExLoop_H
                Call tmpPlane.SetCustomPMD((.Right + 1 + CM_H * (i - 1)), .Top, CM_H, .height)
                Call tmpPlane.CopyPlane(exPlane, pColor)
            Next i
        End If
        If ExBit_V <> 0 Then
             '----- Top Copy -----
            Call exPlane.SetCustomPMD(.Left, .Top, .width, CM_V)
            For i = 1 To ExLoop_V
                Call tmpPlane.SetCustomPMD(.Left, (.Top - CM_V * i), .width, CM_V)
                Call tmpPlane.CopyPlane(exPlane, pColor)
            Next i
            
            '----- Bottom Copy -----
            Call exPlane.SetCustomPMD(.Left, (.Bottom - CM_V + 1), .width, CM_V)
            For i = 1 To ExLoop_V
                Call tmpPlane.SetCustomPMD(.Left, (.Bottom + 1 + CM_V * (i - 1)), .width, CM_V)
                Call tmpPlane.CopyPlane(exPlane, pColor)
            Next i
        End If

        '-------- MEDIAN FILTER ---------
        Call exPlane.SetCustomPMD((.Left - ExBit_H), (.Top - ExBit_V), (.width + 2 * ExBit_H), (.height + 2 * ExBit_V))
        Call tmpPlane.SetCustomPMD((.Left - ExBit_H), (.Top - ExBit_V), (.width + 2 * ExBit_H), (.height + 2 * ExBit_V))
        Call exPlane.RankFilter(tmpPlane, medianTap_H, medianTap_V, ((medianTap_H + medianTap_V) / 2), pColor)

        '-------- exPlane -> srcPlane ---------
        Call exPlane.SetCustomPMD(.Left, .Top, .width, .height)
        Call dstPlane.SetPMD(Zone)
        Call dstPlane.CopyPlane(exPlane, pColor)
    End With
    
End Sub

Public Sub GetLineAddress(ByRef srcPlane As CImgPlane, ByVal Zone As String, ByRef OutVar() As Double, ByRef LineLevel As Variant, ByVal LineType As String, ByVal countType As String)
    
    Dim site As Long
    Dim tmp_Count(nSite) As Double
    
    Select Case countType
        Case "MAX"
            Call Count(srcPlane, Zone, EEE_COLOR_FLAT, idpCountAbove, LineLevel, LineLevel, idpLimitInclude, tmp_Count, "LineDef")
        Case "ABSMAX"
            Call Count(srcPlane, Zone, EEE_COLOR_FLAT, idpCountBetween, LineLevel, LineLevel, idpLimitInclude, tmp_Count, "LineDef")
        Case "MIN"
            Call Count(srcPlane, Zone, EEE_COLOR_FLAT, idpCountBelow, LineLevel, LineLevel, idpLimitInclude, tmp_Count, "LineDef")
        Case Else
            MsgBox "SelectError! @GetLineAddress"
    End Select
        
    Dim PixelLogResult() As T_PIXINFO
    Dim x As Long, y As Long, Data As Double

    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
        
            If tmp_Count(site) <= 0 Then GoTo Next_Site4
            
            With srcPlane
                Call .SetPMD(Zone)
                Call .PixelLog(site, "LineDef", PixelLogResult, 1, idpAddrAbsolute)
            End With

            If LineType = "HLINE" Then
                OutVar(site) = PixelLogResult(0).y
            ElseIf LineType = "VLINE" Then
                OutVar(site) = PixelLogResult(0).x
            Else
                MsgBox "SelectError! @GetLineAddress"
            End If

        End If

Next_Site4:
    Next site
        
    Call TheIDP.PlaneManager(srcPlane.planeGroup).GetSharedFlagPlane("LineDef").RemoveFlagBit("LineDef")
    
End Sub


Public Sub MakeOrPMD(ByVal pPlaneGroup As String, ByVal pSubPmdName As String, ParamArray Zone() As Variant)

'    Dim pSubPmdName As String
    Dim i As Long
    Dim pX As Long
    Dim pY As Long
    Dim pWidth As Long
    Dim pHeight As Long
    Dim X_Start As Long
    Dim X_End As Long
    Dim Y_Start As Long
    Dim Y_End As Long
    Dim tmp_X_Start As Long
    Dim tmp_X_End As Long
    Dim tmp_Y_Start As Long
    Dim tmp_Y_End As Long
'
'    pSubPmdName = Zone(0)
'    For i = 1 To UBound(Zone)
'        pSubPmdName = pSubPmdName & "_" & Zone(i)
'    Next i
 
    tmp_X_Start = TheIDP.PMD(Zone(0)).Left
    tmp_X_End = TheIDP.PMD(Zone(0)).Right
    tmp_Y_Start = TheIDP.PMD(Zone(0)).Top
    tmp_Y_End = TheIDP.PMD(Zone(0)).Bottom
    
    For i = 1 To UBound(Zone)
        X_Start = TheIDP.PMD(Zone(i)).Left
        X_End = TheIDP.PMD(Zone(i)).Right
        Y_Start = TheIDP.PMD(Zone(i)).Top
        Y_End = TheIDP.PMD(Zone(i)).Bottom
        
        'X_Start = MIN
        If tmp_X_Start > X_Start Then tmp_X_Start = X_Start
        'X_End = MAX
        If tmp_X_End < X_End Then tmp_X_End = X_End
        'Y_Start = MIN
        If tmp_Y_Start > Y_Start Then tmp_Y_Start = Y_Start
        'Y_End = MAX
        If tmp_Y_End < Y_End Then tmp_Y_End = Y_End
    Next i
    
    'Create Zone
    pX = tmp_X_Start
    pY = tmp_Y_Start
    pWidth = tmp_X_End - tmp_X_Start + 1
    pHeight = tmp_Y_End - tmp_Y_Start + 1
    
    If TheIDP.isExistingPMD(pSubPmdName) = False Then
        Call TheIDP.PlaneManager(pPlaneGroup).CreateSubPMD(pSubPmdName, pX, pY, pWidth, pHeight)
    End If
    
End Sub

Public Sub SetCalcPmd(ByVal pSubPmdName As String, ByVal pPlaneGroup As String, _
                    ByVal pX As Long, ByVal pY As Long, ByVal pWidth As Long, ByVal pHeight As Long)

    If TheIDP.isExistingPMD(pSubPmdName) = False Then
        Call TheIDP.PlaneManager(pPlaneGroup).CreateSubPMD(pSubPmdName, pX, pY, pWidth, pHeight)
    End If
    
End Sub

Public Function CalcZoneSize(ByVal ScrZoneSize As Double, ByVal CM_Size As Double, ByVal comp As Long) As Double

    Dim Fix_S As Double
    Dim Mod_S As Double

    Fix_S = Int(Int(ScrZoneSize / CM_Size) / comp) * CM_Size
    Mod_S = ScrZoneSize - (comp * Fix_S)
    If Mod_S < CM_Size Then
        CalcZoneSize = Fix_S + Mod_S
    Else
        CalcZoneSize = Fix_S + CM_Size
    End If
    
End Function

Private Function DoubleToArray(res As Variant) As Double()
    Dim site As Long
    Dim t_result(nSite) As Double
    
    If IsArray(res) Then
        DoubleToArray = res
    Else
    
        For site = 0 To nSite
            t_result(site) = res
        Next site
        DoubleToArray = t_result
    End If
End Function

Public Function MakeSliceLevel(ByRef OutVar() As Double, ByVal Inslice As Variant, Optional ByVal Lsb As Variant = 1, Optional ByVal BaseLevel As Variant = Null, Optional ByVal operation As String = "", Optional ByVal Kco As Double = 1, Optional ByVal countType As String = "0")

    Dim site As Long
    ReDim InsliceCon(nSite) As Double
    InsliceCon = DoubleToArray(Inslice)
    
    If IsNull(BaseLevel) = True Then
        For site = 0 To nSite
            If IsArray(Lsb) Then
                OutVar(site) = InsliceCon(site) / Lsb(site) / Kco
            Else
                OutVar(site) = InsliceCon(site) / Kco
            End If
            If countType = 2 Then OutVar(site) = Int(OutVar(site)) - 1
            If countType = 1 Then OutVar(site) = Int(OutVar(site)) + 2
        Next site
    Else
        For site = 0 To nSite
            If IsArray(Lsb) Then
                OutVar(site) = InsliceCon(site) / Lsb(site) / Kco
            Else
                OutVar(site) = InsliceCon(site) / Kco
            End If
            Select Case operation
                Case "+"
                    OutVar(site) = OutVar(site) + BaseLevel(site)
                Case "-"
                    OutVar(site) = OutVar(site) - BaseLevel(site)
                Case "*"
                    OutVar(site) = OutVar(site) * BaseLevel(site)
                Case "/"
                    OutVar(site) = Div(OutVar(site), BaseLevel(site), 0)
            End Select

            If countType = 2 Then OutVar(site) = Int(OutVar(site)) - 1
            If countType = 1 Then OutVar(site) = Int(OutVar(site)) + 2
        Next site
    End If

End Function

Public Function MakeSliceLevel_Percent(ByRef OutVar() As Double, ByVal Inslice As Variant)

    Dim site As Long
    ReDim InsliceCon(nSite) As Double
    InsliceCon = DoubleToArray(Inslice)

    For site = 0 To nSite
        OutVar(site) = 1 + InsliceCon(site)
    Next site

End Function

Public Function ResultAdd(ByRef TsetLabel As String, ByRef pInput() As Double)

    If TheResult.IsExist(TsetLabel) = True Then
        TheResult.Delete (TsetLabel)
    End If
        
    TheResult.Add TsetLabel, pInput

End Function

Public Function ResultAddEX(ByRef TsetLabel As String, ByRef pInput() As Double)

'TheResult.Addとオフセットの反映まで行う。
'Test InstancesでReturnResultEx_fをするとオフセットが2回かかるので注意が必要

    '既に登録されている場合は削除し登録し直す
    If TheResult.IsExist(TsetLabel) = True Then
        TheResult.Delete (TsetLabel)
    End If
        
    TheResult.Add TsetLabel, pInput

    '@@@ 測定結果コンバート @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    TheOffsetResult.Calculate TsetLabel, TheResult
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

End Function

''For Deffect
Public Function d_read_vmcu_Line(ByRef srcPlane As CImgPlane, _
                                  ByVal Zone As String, _
                                  ByVal DefectNum As Long, _
                                  ByVal Lsb As Variant, _
                                  ByVal Kco As Variant, _
                                  ByVal signature As String, _
                                  ByVal Unit As String, _
                                  ByVal LineLevel As Variant, _
                                  ByVal LineType As String, _
                                  ByVal countType As String, _
                                  Optional ByVal BaseVal As Variant = 1)

    Dim site As Long
    Dim LineAdd(nSite) As Double
    Call GetLineAddress(srcPlane, Zone, LineAdd, LineLevel, LineType, countType)
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim Data As Double
    Dim DefectNumTmp As Long
    DefectNumTmp = 1

    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
        
            Dim currentSiteDeviceNumber As String
            currentSiteDeviceNumber = getActiveSiteDeviceNumer(site)
            
            For i = 0 To DefectNumTmp - 1

                Select Case LineType
                    Case "HLINE"
                        x = 1
                        y = LineAdd(site)
                    Case "VLINE"
                        x = LineAdd(site)
                        y = 1
                    Case Else
                        MsgBox "SelectError! @d_read_vmcu_Line"
                End Select

                If Kco = "NoKCO" Then Kco = 1
                Select Case Unit
                    Case "uV"
                        Data = LineLevel(site) * Lsb(site) * Kco / uV
                    Case "mV"
                        Data = LineLevel(site) * Lsb(site) * Kco / mV
                    Case "V"
                        Data = LineLevel(site) * Lsb(site) * Kco / V
                    Case "%"
                        Data = Div(LineLevel(site), BaseVal(site), 0) * 100
                    Case "LSB"
                        Data = LineLevel(site)
                    Case Else
                        Data = LineLevel(site) * Lsb(site)
                End Select
        
                If Sw_Ana = 1 Then
                    TheExec.Datalog.WriteComment "*****  DEFECT " & LineType & " DATA (SITE:" & site & ") *****"
                    Print #m_fileLunDefectFile, returnCurrentDefectLine(currentSiteDeviceNumber, signature, Unit, x, y, Data, True)
                End If

                If Flg_Debug = 1 Then
                    TheExec.Datalog.WriteComment "*****  DEFECT " & LineType & " DATA (SITE:" & site & ") *****"
                    TheExec.Datalog.WriteComment returnCurrentDefectLine(currentSiteDeviceNumber, signature, Unit, x, y, Data, False)
                End If
            Next i
        End If
    Next site

End Function

Public Function d_read_vmcu_Point(ByRef srcPlane As CImgPlane, _
                                  ByVal Zone As String, _
                                  ByVal MaxCount As Double, _
                                  ByVal Lsb As Variant, _
                                  ByVal Kco As Variant, _
                                  ByVal signature As String, _
                                  ByVal Unit As String, _
                                  ByVal countType As String, _
                                  ByVal PointLevel As Variant, _
                                  ByVal CountLimit As String, _
                                  ByVal InputFlg As String, _
                                  ParamArray ColorArr() As Variant)

    Dim site As Long
    Dim tmp_count_All As CImgColorAllResult
    Dim tmp_Count(nSite) As Double
    Dim tmp_count_color(nSite) As Double
    Dim ColorMap() As String
    Dim PixelLogResult() As T_PIXINFO
    Dim x As Long, y As Long, Data As Double
    Dim currentSiteDeviceNumber As String
    Dim i As Long
    Dim j As Long

    If UBound(ColorArr) = -1 Then
        MsgBox "SelectError! @d_read_vmcu_Point"
    End If

    ColorMap = srcPlane.planeMap.ColorList

    If UBound(ColorMap) = UBound(ColorArr) Or ColorArr(0) = "-" Then
        If InputFlg <> "NoInputFlg" Then
            Call CountColorAll(srcPlane, Zone, countType, PointLevel, PointLevel, idpLimitEachSite, CountLimit, tmp_count_All, "PointDef", InputFlg)
        Else
            Call CountColorAll(srcPlane, Zone, countType, PointLevel, PointLevel, idpLimitEachSite, CountLimit, tmp_count_All, "PointDef")
        End If
        Call GetSum_Color_Flat(tmp_Count, tmp_count_All)
            
        For site = 0 To nSite
            If tmp_Count(site) < MaxCount Then
                tmp_Count(site) = tmp_Count(site)
            Else
                tmp_Count(site) = MaxCount
            End If
        Next site
        
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
            
                If tmp_Count(site) <= 0 Then GoTo Next_Site
                currentSiteDeviceNumber = getActiveSiteDeviceNumer(site)
                
                ReDim PixelLogResult(tmp_Count(site))
                
                With srcPlane
                    Call .SetPMD(Zone)
                    Call .PixelLog(site, "PointDef", PixelLogResult, tmp_Count(site), idpAddrAbsolute)
                End With
    
                If Sw_Ana = 1 Or Flg_Debug = 1 Then
                    TheExec.Datalog.WriteComment "*****  DEFECT " & signature & " DATA (SITE:" & site & ") *****"
                End If
    
                For i = 0 To tmp_Count(site) - 1
                    x = PixelLogResult(i).x
                    y = PixelLogResult(i).y
                    Data = PixelLogResult(i).Value

                    If Kco = "NoKCO" Then Kco = 1
                    Select Case Unit
                        Case "uV"
                            Data = Data * Lsb(site) * Kco / uV
                        Case "mV"
                            Data = Data * Lsb(site) * Kco / mV
                        Case "V"
                            Data = Data * Lsb(site) * Kco / V
                        Case "%"
                            Data = (Data - 1) * 100
                        Case Else
                            Data = Data * Lsb(site)
                    End Select
                         
                    If Sw_Ana = 1 Then
                        Print #m_fileLunDefectFile, returnCurrentDefectLine(currentSiteDeviceNumber, signature, Unit, x, y, Data, True)
                    End If
                    If Flg_Debug = 1 Then
                        TheExec.Datalog.WriteComment returnCurrentDefectLine(currentSiteDeviceNumber, signature, Unit, x, y, Data, False)
                    End If
                Next i
            End If
Next_Site:
        Next site
            
        Call srcPlane.GetSharedFlagPlane("PointDef").RemoveFlagBit("PointDef")
            
    Else
            
        For i = 0 To UBound(ColorArr)
        
            If InputFlg <> "NoInputFlg" Then
                Call Count(srcPlane, Zone, ColorArr(i), countType, PointLevel, PointLevel, CountLimit, tmp_Count, "PointDef", InputFlg)
            Else
                Call Count(srcPlane, Zone, ColorArr(i), countType, PointLevel, PointLevel, CountLimit, tmp_Count, "PointDef")
            End If
            
            For site = 0 To nSite
                tmp_Count(site) = tmp_Count(site)
            Next site
            
            For site = 0 To nSite
                If tmp_Count(site) < MaxCount Then
                    tmp_Count(site) = tmp_Count(site)
                Else
                    tmp_Count(site) = MaxCount
                End If
            Next site
            
            For site = 0 To nSite
                If TheExec.sites.site(site).Active Then

                    If tmp_Count(site) <= 0 Then GoTo Next_Site2
                    
                    currentSiteDeviceNumber = getActiveSiteDeviceNumer(site)
                    
                    ReDim PixelLogResult(tmp_Count(site))
                    
                    With srcPlane
                        Call .SetPMD(Zone)
                        Call .PixelLog(site, "PointDef", PixelLogResult, tmp_Count(site), idpAddrAbsolute)
                    End With
        
                    If Sw_Ana = 1 Or Flg_Debug = 1 Then
                        TheExec.Datalog.WriteComment "*****  DEFECT " & signature & " DATA (SITE:" & site & " , COLOR:" & ColorArr(i) & ") *****"
                    End If
        
                    For j = 0 To tmp_Count(site) - 1
                        x = PixelLogResult(j).x
                        y = PixelLogResult(j).y
                        Data = PixelLogResult(j).Value
                        
                        If Kco = "NoKCO" Then Kco = 1
                        Select Case Unit
                            Case "uV"
                                Data = Data * Lsb(site) * Kco / uV
                            Case "mV"
                                Data = Data * Lsb(site) * Kco / mV
                            Case "V"
                                Data = Data * Lsb(site) * Kco / V
                            Case "%"
                                Data = (Data - 1) * 100
                            Case Else
                                Data = Data * Lsb(site)
                        End Select
                             
                        If Sw_Ana = 1 Then
                            Print #m_fileLunDefectFile, returnCurrentDefectLine(currentSiteDeviceNumber, signature, Unit, x, y, Data, True)
                        End If
                        If Flg_Debug = 1 Then
                            TheExec.Datalog.WriteComment returnCurrentDefectLine(currentSiteDeviceNumber, signature, Unit, x, y, Data, False)
                        End If
                    Next j
                End If
Next_Site2:
            Next site
            
        Call srcPlane.GetSharedFlagPlane("PointDef").RemoveFlagBit("PointDef")
                                    
        Next i
    End If

End Function

' 内部でPutFlagを行う（Plane枚数が一枚多い＋時間がかかる＋フラグBitを一ビット使用）
Public Function d_read_vmcu_FD(ByRef srcPlane As CImgPlane, _
                                    ByRef coutPlane As CImgPlane, _
                                    ByVal Zone As String, _
                                    ByVal MaxCount As Double, _
                                    ByVal Lsb As Variant, _
                                    ByVal Kco As Variant, _
                                    ByVal signature As String, _
                                    ByVal Unit As String, ByVal tmp_Count As Variant, ByVal Inflagname As String, _
                                    ByVal offsetX As Double, ByVal offsetY As Double, _
                                    ByVal MulX As Double, ByVal MulY As Double)
                                    
    Dim site As Long
    Dim PixelLogResult() As T_PIXINFO
    Dim x As Long, y As Long, Data As Double
    Dim i As Long
    Dim currentSiteDeviceNumber As String
    Dim Color_count As Double
    
    Color_count = srcPlane.planeMap.Count
    
    
    For site = 0 To nSite
        tmp_Count(site) = tmp_Count(site) * Color_count
        If tmp_Count(site) < MaxCount Then
            tmp_Count(site) = tmp_Count(site)
        Else
            tmp_Count(site) = MaxCount
        End If
    Next site
    
    Dim tmpPix(nSite) As CPixInfo
    Call ReadPixelSite(coutPlane, Zone, tmp_Count, Inflagname, tmpPix, idpAddrAbsolute)

    Dim tmpPix2(nSite) As CPixInfo
    Call RPDOffset(tmpPix2, tmpPix, offsetX, offsetY, MulX, MulY, 1)

    Dim tmpPix3(nSite) As CPixInfo
    
    Dim xx As Double
    Dim yy As Double

    For yy = 0 To MulY - 1
        For xx = 0 To MulX - 1
            Call RPDOffset(tmpPix, tmpPix2, xx, yy)
            Call RPDUnion(tmpPix3, tmpPix3, tmpPix)
        Next xx
    Next yy

    Dim tmpPlane As CImgPlane
    Call GetFreePlane(tmpPlane, srcPlane.planeGroup, idpDepthS16, True)
    Call WritePixelAddrSite(tmpPlane, tmpPlane.BasePMD, tmpPix3)

    Call PutFlag(tmpPlane, tmpPlane.BasePMD, EEE_COLOR_ALL, idpCountAbove, 1, 1, idpLimitInclude, "FLG_FD_Defect")
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
        
            If tmpPix3(site).Count <= 0 Then GoTo NextSite

            currentSiteDeviceNumber = getActiveSiteDeviceNumer(site)
            
            ReDim PixelLogResult(tmp_Count(site))
            
            
            With srcPlane
'                Call .SetPMD(Zone)
                Call .SetPMD(.BasePMD)
                Call .PixelLog(site, "FLG_FD_Defect", PixelLogResult, tmp_Count(site), idpAddrAbsolute)
            End With
            
            If Sw_Ana = 1 Or Flg_Debug = 1 Then
                TheExec.Datalog.WriteComment "*****  DEFECT " & signature & " DATA (SITE:" & site & ") *****"
            End If

            For i = 0 To tmp_Count(site) - 1
                x = PixelLogResult(i).x
                y = PixelLogResult(i).y
                Data = PixelLogResult(i).Value

                If Kco = "NoKCO" Then Kco = 1
                Select Case Unit
                    Case "uV"
                        Data = Data * Lsb(site) * Kco / uV
                    Case "mV"
                        Data = Data * Lsb(site) * Kco / mV
                    Case "V"
                        Data = Data * Lsb(site) * Kco / V
                    Case "%"
                        Data = (Data - 1) * 100
                    Case Else
                        Data = Data * Lsb(site)
                End Select
                     
                If Sw_Ana = 1 Then
                    Print #m_fileLunDefectFile, returnCurrentDefectLine(currentSiteDeviceNumber, signature, Unit, x, y, Data, True)
                End If
                If Flg_Debug = 1 Then
                    TheExec.Datalog.WriteComment returnCurrentDefectLine(currentSiteDeviceNumber, signature, Unit, x, y, Data, False)
                End If
                
            Next i
        End If
NextSite:
    Next site
    
    Call srcPlane.GetSharedFlagPlane("FLG_FD_Defect").RemoveFlagBit("FLG_FD_Defect")

End Function

Public Function d_read_vmcu_Point_FD(ByRef srcPlane As CImgPlane, _
                                    ByVal Zone As String, _
                                    ByVal MaxCount As Double, _
                                    ByVal Lsb As Variant, _
                                    ByVal Kco As Variant, _
                                    ByVal signature As String, _
                                    ByVal Unit As String, ByVal tmp_Count As Variant, ByVal Inflagname As String)
                                    
    Dim site As Long
    Dim PixelLogResult() As T_PIXINFO
    Dim x As Long, y As Long, Data As Double
    Dim i As Long
    Dim currentSiteDeviceNumber As String
    Dim Color_count As Double
    
    Color_count = srcPlane.planeMap.Count
    
    For site = 0 To nSite
        tmp_Count(site) = tmp_Count(site) * Color_count
        If tmp_Count(site) < MaxCount Then
            tmp_Count(site) = tmp_Count(site)
        Else
            tmp_Count(site) = MaxCount
        End If
    Next site
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
        
            If tmp_Count(site) <= 0 Then GoTo NextSite

            currentSiteDeviceNumber = getActiveSiteDeviceNumer(site)
            
            ReDim PixelLogResult(tmp_Count(site))
            
            With srcPlane
                Call .SetPMD(Zone)
                Call .PixelLog(site, Inflagname, PixelLogResult, tmp_Count(site), idpAddrAbsolute)
            End With
            
            If Sw_Ana = 1 Or Flg_Debug = 1 Then
                TheExec.Datalog.WriteComment "*****  DEFECT " & signature & " DATA (SITE:" & site & ") *****"
            End If

            For i = 0 To tmp_Count(site) - 1
                x = PixelLogResult(i).x
                y = PixelLogResult(i).y
                Data = PixelLogResult(i).Value

                If Kco = "NoKCO" Then Kco = 1
                Select Case Unit
                    Case "uV"
                        Data = Data * Lsb(site) * Kco / uV
                    Case "mV"
                        Data = Data * Lsb(site) * Kco / mV
                    Case "V"
                        Data = Data * Lsb(site) * Kco / V
                    Case "%"
                        Data = (Data - 1) * 100
                    Case Else
                        Data = Data * Lsb(site)
                End Select
                     
                If Sw_Ana = 1 Then
                    Print #m_fileLunDefectFile, returnCurrentDefectLine(currentSiteDeviceNumber, signature, Unit, x, y, Data, True)
                End If
                If Flg_Debug = 1 Then
                    TheExec.Datalog.WriteComment returnCurrentDefectLine(currentSiteDeviceNumber, signature, Unit, x, y, Data, False)
                End If
                
            Next i
        End If
NextSite:
    Next site
        
    Call srcPlane.GetSharedFlagPlane("PointDef").RemoveFlagBit("PointDef")

End Function

Private Function returnCurrentDefectLine( _
    ByVal currentDeviceNumber As String, _
    ByVal defectLabel As String, _
    ByVal defectUnit As String, _
    ByVal xAddress As Long, _
    ByVal yAddress As Long, _
    ByVal writeData As Double, _
    Optional ByVal isTargetFile As Boolean = True, _
    Optional ByVal dataMultiplyer As Long = 1000, _
    Optional ByVal defectDeliminator As String = " " _
    ) As String
    
    If isTargetFile Then
        returnCurrentDefectLine = currentDeviceNumber & defectDeliminator _
                                & defectLabel & defectDeliminator _
                                & defectUnit & defectDeliminator _
                                & Format(xAddress, "####") & defectDeliminator _
                                & Format(yAddress, "####") & defectDeliminator _
                                & Format(writeData * dataMultiplyer, "######")
    Else
        returnCurrentDefectLine = defectLabel & defectDeliminator _
                                & ":(" _
                                & Format(xAddress, "####") _
                                & ", " _
                                & Format(yAddress, "####") _
                                & ") = " _
                                & Format(writeData, "##0.##0") _
                                & " " _
                                & defectUnit
    End If

End Function

Private Sub GetCurrentPmdInfo(ByRef srcPlane As CImgPlane, ByRef pPmdInfo As pmdInfo)
    With srcPlane.CurrentPMD
        pPmdInfo.Left = .Left
        pPmdInfo.Top = .Top
        pPmdInfo.width = .width
        pPmdInfo.height = .height
        pPmdInfo.Right = .Right
        pPmdInfo.Bottom = .Bottom
    End With
End Sub




''APPENDIX    notUSE
Public Sub MakeMulPMD(ByRef srcPlane As CImgPlane, ByVal Zone As String, ByVal pSubPmdName As String, _
                    ByVal comp_X As Long, ByVal comp_Y As Long, ByVal pColor As String)

    Dim pX As Long
    Dim pY As Long
    Dim pWidth As Long
    Dim pHeight As Long
    Dim tmp_X_Width As Long
    Dim tmp_Y_Height As Long
    Dim CM_X As Double
    Dim CM_Y As Double
    Dim Mod_X As Double
    Dim Mod_Y As Double
    Dim Fix_X As Long
    Dim Fix_Y As Long

    If TheIDP.isExistingPMD(pSubPmdName) Then Exit Sub
    
    tmp_X_Width = TheIDP.PMD(Zone).width
    tmp_Y_Height = TheIDP.PMD(Zone).height
    
    If pColor = EEE_COLOR_ALL Then
        CM_X = srcPlane.planeMap.width
        CM_Y = srcPlane.planeMap.height
    End If
    
    If pColor = EEE_COLOR_FLAT Then
        CM_X = 1
        CM_Y = 1
    End If

    ' pX = 1
    ' pY = 1
    
    ' 左上に詰める
    pX = Int((TheIDP.PMD(Zone).XAdr - 1) / comp_X) + 1
    pY = Int((TheIDP.PMD(Zone).YAdr - 1) / comp_Y) + 1
    
    If CM_X > tmp_X_Width Or CM_Y > tmp_Y_Height Then
        MsgBox "SelectError! @MakeMulPMD"
    End If

    ''H方向
    Fix_X = Int(Int(tmp_X_Width / CM_X) / comp_X) * CM_X
    Mod_X = tmp_X_Width - (comp_X * Fix_X)
    If Mod_X < CM_X Then
         pWidth = Fix_X + Mod_X
    Else
     pWidth = Fix_X + CM_X
    End If

    ''V方向
    Fix_Y = Int(Int(tmp_Y_Height / CM_Y) / comp_Y) * CM_Y
    Mod_Y = tmp_Y_Height - (comp_Y * Fix_Y)
    If Mod_Y < CM_Y Then
         pHeight = Fix_Y + Mod_Y
    Else
     pHeight = Fix_Y + CM_Y
    End If

    Call TheIDP.PlaneManager(srcPlane.planeGroup).CreateSubPMD(pSubPmdName, pX, pY, pWidth, pHeight)
    
End Sub

Public Sub MakeAccPMD(ByRef srcPlane As CImgPlane, ByVal Zone As String, ByVal pSubPmdName As String, ByVal pColor As String)
    
    Dim pX As Long
    Dim pY As Long
    Dim pWidth As Long
    Dim pHeight As Long
    
    pX = TheIDP.PMD(Zone).Left
    pY = 1
    
    pWidth = TheIDP.PMD(Zone).width
    If pColor = EEE_COLOR_ALL Then pHeight = srcPlane.planeMap.height
    If pColor = EEE_COLOR_FLAT Then pHeight = 1
    
    If TheIDP.isExistingPMD(pSubPmdName) = False Then
        Call TheIDP.PlaneManager(srcPlane.planeGroup).CreateSubPMD(pSubPmdName, pX, pY, pWidth, pHeight)
    End If
    
End Sub

Public Sub MakeAccJudgePMD(ByRef srcPlane As CImgPlane, ByVal Zone As String, ByVal pSubPmdName As String, ByVal pColor As String, ByVal Diff As Long)
    
    Dim pX As Long
    Dim pY As Long
    Dim pWidth As Long
    Dim pHeight As Long
    
    pX = TheIDP.PMD(Zone).Left
    pY = 1
    
    pWidth = TheIDP.PMD(Zone).width - Diff
    If pColor = EEE_COLOR_ALL Then pHeight = srcPlane.planeMap.height
    If pColor = EEE_COLOR_FLAT Then pHeight = 1
    
    If TheIDP.isExistingPMD(pSubPmdName) = False Then
        Call TheIDP.PlaneManager(srcPlane.planeGroup).CreateSubPMD(pSubPmdName, pX, pY, pWidth, pHeight)
    End If
    
End Sub

Public Sub MakeAcrPMD(ByRef srcPlane As CImgPlane, ByVal Zone As String, ByVal pSubPmdName As String, ByVal pColor As String)
    
    Dim pX As Long
    Dim pY As Long
    Dim pWidth As Long
    Dim pHeight As Long
    
    pX = 1
    pY = TheIDP.PMD(Zone).Top
    
    If pColor = EEE_COLOR_ALL Then pWidth = srcPlane.planeMap.width
    If pColor = EEE_COLOR_FLAT Then pWidth = 1
    pHeight = TheIDP.PMD(Zone).height
    
    If TheIDP.isExistingPMD(pSubPmdName) = False Then
        Call TheIDP.PlaneManager(srcPlane.planeGroup).CreateSubPMD(pSubPmdName, pX, pY, pWidth, pHeight)
    End If
    
End Sub

Public Sub MakeAcrJudgePMD(ByRef srcPlane As CImgPlane, ByVal Zone As String, ByVal pSubPmdName As String, ByVal pColor As String, ByVal Diff As Long)
    
    Dim pX As Long
    Dim pY As Long
    Dim pWidth As Long
    Dim pHeight As Long
    
    pX = 1
    pY = TheIDP.PMD(Zone).Top
    If pColor = EEE_COLOR_ALL Then pWidth = srcPlane.planeMap.width
    If pColor = EEE_COLOR_FLAT Then pWidth = 1
    pHeight = TheIDP.PMD(Zone).height - Diff
    
    If TheIDP.isExistingPMD(pSubPmdName) = False Then
        Call TheIDP.PlaneManager(srcPlane.planeGroup).CreateSubPMD(pSubPmdName, pX, pY, pWidth, pHeight)
    End If
    
End Sub

Public Function GetSumMax(ParamArray dataArr() As Variant) As Double

    Dim site As Long
    Dim i As Long
    Dim tmpSum(nSite) As Double
    Dim tmpMax As Double
    
    tmpMax = dataArr(0)(0)
    For site = 0 To nSite
        tmpSum(site) = dataArr(0)(site)
        For i = 1 To UBound(dataArr)
            tmpSum(site) = tmpSum(site) + dataArr(i)(site)
        Next i
        
        If tmpMax < tmpSum(site) Then
            tmpMax = tmpSum(site)
        End If
        
    Next site
    
    GetSumMax = tmpMax

End Function

Public Function SumVariable(ParamArray dataArr() As Variant) As Double
 '三池確認必要

    Dim site As Long
    Dim i As Long
    Dim tmp As Double
        
    For site = 0 To nSite
        For i = 0 To UBound(dataArr)
            tmp = tmp + dataArr(i)(site)
        Next i
    Next site
    SumVariable = tmp

End Function

Public Function GetSum_Color_Flat(ByRef OutVar() As Double, ByRef inVar As CImgColorAllResult)

    Dim i As Long
    Dim site As Long
    Dim tmp(nSite) As Double
    Dim ColorMap() As String
    
'    ColorMap = tmp1.ColorList
    ColorMap = inVar.ColorList
    For site = 0 To nSite
        tmp(site) = inVar.color(ColorMap(0)).SiteValue(site)
        For i = 1 To UBound(ColorMap)
            tmp(site) = tmp(site) + inVar.color(ColorMap(i)).SiteValue(site)
        Next i
        OutVar(site) = tmp(site)
    Next site

End Function

Private Function getActiveSiteDeviceNumer(ByVal currentSite As Long) As String

    Const LenChipNumber As Long = 4
    Dim currentWaferNumber As String
    Dim formatString As String
    Dim i As Long
    
    If Sw_Ana = 1 Then
        formatString = ""
        For i = 0 To LenChipNumber - 1
            formatString = formatString & "0"
        Next i
        
        currentWaferNumber = Mid(CStr(DeviceNumber), 1, Len(CStr(DeviceNumber)) - LenChipNumber)
        getActiveSiteDeviceNumer = currentWaferNumber & Format(CStr(DeviceNumber_site(currentSite)), formatString)
    End If
    
End Function

Public Sub OutPutImage( _
    ByVal site As Long, ByRef plane As CImgPlane, ByVal Zone As Variant, ByVal fileName As String, _
    Optional fileType As IdpFileFormat = idpFileBinary)
    
    On Error GoTo ErrorEnd
    
    Call plane.SetPMD(Zone)
    Call plane.WriteFile(site, fileName, fileType)

ErrorEnd:

End Sub

Public Sub InPutImage( _
    ByVal site As Long, ByRef plane As CImgPlane, ByVal Zone As Variant, ByVal fileName As String, _
    Optional fileType As IdpFileFormat = idpFileBinary)
    
    Call plane.SetPMD(Zone)
    Call plane.ReadFile(site, fileName, fileType)

    TheExec.Datalog.WriteComment "Input IMAGE DATA!! " & fileName & ""

End Sub

Public Sub MedianEx_NonContinuous(ByRef srcPlane As CImgPlane, ByRef dstPlane As CImgPlane, ByVal Zone As Variant, ByVal medianTap_H As Double, ByVal medianTap_V As Double, _
                                  ByVal CycleDirection As String, ByVal CycleSize As Double, ByVal CycleCount As Double)

    'エラー処理
    If medianTap_H <> 1 And medianTap_V <> 1 Then
        MsgBox "SelectError! @MedianEx_NonContinuous"
    End If

    Dim tmp_H As Double
    Dim tmp_V As Double
    Dim Center As Double
    Dim CycleH As Double
    Dim CycleV As Double
    
    If CycleDirection = "V" Then
        CycleV = CycleSize * CycleCount
        CycleH = srcPlane.planeMap.width
    ElseIf CycleDirection = "H" Then
        CycleV = srcPlane.planeMap.height
        CycleH = CycleSize * CycleCount
    End If
    
    tmp_H = 2 / 3 * TheIDP.PMD(Zone).width / CycleH + 1
    tmp_V = 2 / 3 * TheIDP.PMD(Zone).height / CycleV + 1
    
    If tmp_H <= medianTap_H Then
        If Int(tmp_H) Mod 2 = 1 Then
            medianTap_H = Int(tmp_H)
        Else
            medianTap_H = Int(tmp_H) - 1
        End If
    End If
    If tmp_V <= medianTap_V Then
        If Int(tmp_V) Mod 2 = 1 Then
            medianTap_V = Int(tmp_V)
        Else
            medianTap_V = Int(tmp_V) - 1
        End If
    End If
        
    Center = Int((medianTap_H * medianTap_V + 1) / 2)
    Call srcPlane.SetPMD(Zone)
    Call dstPlane.SetPMD(Zone)
    Call dstPlane.NonContinuousRankFilter(srcPlane, medianTap_H, medianTap_V, Center, CycleH, CycleV)
    Call Extention(dstPlane, Zone, dstPlane, -Int(medianTap_H / 2) * CycleH, -Int(medianTap_H / 2) * CycleH, -Int(medianTap_V / 2) * CycleV, -Int(medianTap_V / 2) * CycleV, EEE_COLOR_FLAT)

End Sub

Public Sub Make_Input_Flag(ByRef srcPlane As CImgPlane, ByVal Direcrion As String, ByVal Start_Add As Long, ByVal Cycle As Long, ParamArray FlgNameArr() As Variant)

'Updated by S.Matsuno 2013-11-14 (Master book installed by T.Morimoto 2013-11-18)

    Dim site As Long
    Dim pX As Long, pY As Long, pWidth As Long, pHeight As Long
    Dim EachFLGName As String, outputPlaneName As String
    
    Dim modPMD As Long
    Dim n As Long, Cycle_Count As Integer, Cycle_End As Integer
    Dim Flg_ActiveSite(nSite) As Long
    Dim Flg_count As Long
    Flg_count = UBound(FlgNameArr) + 1
    
    'SITE ACTIVE Save
    If Flg_FirstCompleteRun = False Then
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                Flg_ActiveSite(site) = 1
            End If
        Next site
        TheExec.sites.SetAllActive (True)
    End If
    
    ' 結果プレーンの確保
    Dim tmpPlane As CImgPlane
    
    EachFLGName = srcPlane.planeGroup & "-" & Direcrion & "SA_" & Start_Add & "Bit" & Cycle & "_C" & Flg_count
    outputPlaneName = "INPUT_FLAG_" & EachFLGName
            
    '' 画像が存在するかチェック
    If TheIDP.PlaneBank.isExisting(outputPlaneName) Then
        ' 画像が存在した場合、画像を読み込み、フラグを作成する。
        Call GetRegisteredPlane(outputPlaneName, tmpPlane, "InputFlagImage")
    Else
        ' 画像が存在しない場合、フラグを作成するための画像を作成し、登録し、フラグを作成する。
        'StartCheck
        Call GetFreePlane(tmpPlane, srcPlane.planeGroup, srcPlane.BitDepth, True, "Input_Flag_Image")
        pX = 1
        pY = 1
        pWidth = tmpPlane.BasePMD.width
        pHeight = tmpPlane.BasePMD.height
        If Start_Add <> 1 Then
            ' 周期開始前のエリアは、最後の周期のBitを立てる(0のほうがよいのでは？)
            If Direcrion = "V" Then pHeight = Start_Add - 1
            If Direcrion = "H" Then pWidth = Start_Add - 1
            Call tmpPlane.SetCustomPMD(pX, pY, pWidth, pHeight)
            Call tmpPlane.WritePixel(Flg_count + 10, EEE_COLOR_ALL)
        End If
        
        If Direcrion = "V" Then Cycle_End = tmpPlane.BasePMD.height - Cycle
        If Direcrion = "H" Then Cycle_End = tmpPlane.BasePMD.width - Cycle
        
        If Start_Add > 1 Then
            n = (Start_Add / Cycle) Mod Flg_count

            If Direcrion = "V" Then
                pY = 1
                pHeight = Start_Add - 1
            ElseIf Direcrion = "H" Then
                pX = 1
                pWidth = Start_Add - 1
            End If
            
            Call tmpPlane.SetCustomPMD(pX, pY, pWidth, pHeight)
            Call tmpPlane.WritePixel(n + 10, EEE_COLOR_ALL)
        End If
        
        For Cycle_Count = Start_Add To Cycle_End Step Cycle
            n = (Cycle_Count - Start_Add) / Cycle Mod Flg_count
                
            If Direcrion = "V" Then
                pY = Cycle_Count
                pHeight = Cycle
            ElseIf Direcrion = "H" Then
                pX = Cycle_Count
                pWidth = Cycle
            End If
            
            Call tmpPlane.SetCustomPMD(pX, pY, pWidth, pHeight)
            Call tmpPlane.WritePixel(n + 10, EEE_COLOR_ALL)
        Next Cycle_Count
        
        ' 余り発生分
        If Cycle_Count <> Cycle_End Then
            n = (Cycle_Count - Start_Add) / Cycle Mod Flg_count
            If Direcrion = "V" Then
                pY = Cycle_Count
                pHeight = tmpPlane.BasePMD.height - pY + 1
            ElseIf Direcrion = "H" Then
                pX = Cycle_Count
                pWidth = tmpPlane.BasePMD.width - pX + 1
            End If
            Call tmpPlane.SetCustomPMD(pX, pY, pWidth, pHeight)
            Call tmpPlane.WritePixel(n + 10, EEE_COLOR_ALL)
        End If
        
        
        ' 作成した画像を登録する
        Call TheIDP.PlaneBank.Add(outputPlaneName, tmpPlane, True, True)
    End If
    
    'Create Flg
    For n = 0 To Flg_count - 1
        Call PutFlag(tmpPlane, tmpPlane.BasePMD.Name, EEE_COLOR_ALL, idpCountBetween, n + 10, n + 10, idpLimitInclude, FlgNameArr(n))
    Next n
        
    'SITE ACTIVE RETURN
    If Flg_FirstCompleteRun = False Then
        For site = 0 To nSite
            If Flg_ActiveSite(site) = 0 Then
                TheExec.sites.site(site).Active = False
            End If
        Next site
    End If
        
End Sub

Public Sub countBitMask( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByVal srcColor As Variant, ByVal countType As IdpCountType, _
    ByVal loLim As Variant, ByVal hiLim As Variant, ByVal limitType As IdpLimitType, ByRef retResult() As Double, _
    Optional ByVal pFlgName As String = "", Optional ByVal pInputFlgName As String = "", Optional ByVal pMaskBit As Long _
)
'内容:
'   条件に該当する点の個数を取得する。
'
'[srcPlane]         IN   CImgPlane型:    対象プレーン
'[srcZone]          IN   String型:       対象プレーンのゾーン指定
'[srcColor]         IN   IdpColorType型: 対象プレーンの色指定
'[countType]        IN   IdpCountType型: カウント条件指定
'[loLim]            IN   Variant型:      下限値
'[hiLim]            IN   Variant型:      上限値
'[limitType]        IN   IdpLimitType型: 境界値を含む、含まない指定
'[retResult()]      OUT  Double型:       結果格納用配列(動的配列)
'[pFlgName]         IN   String型:       出力フラグ名
'[pInputFlgName]    IN   String型:       入力フラグ名
'
'備考:
'   hiLim,loLimをVariantで定義しているのは、サイト別で境界値が違う場合に対応するため。
'   サイト配列を入れると各サイトごと別々に、定数を入れると全サイトに同じ値が適用される。
'   srcColorにidpColorAllを指定するとAllで処理をしてFlatの結果を返す。
'
    Call srcPlane.SetPMD(srcZone)
    Call srcPlane.Count(retResult, countType, loLim, hiLim, limitType, srcColor, pFlgName, pInputFlgName, pMaskBit)

End Sub

Public Function MakeSliceLevel_ShiroKobu_1Step(ByRef OutVar() As Double, ByVal StartLevel As Double, ByVal EndLevel As Double, ByVal Step As Double, Optional ByVal Lsb As Variant = 1, Optional ByVal Kco As Double = 1)

    Dim site As Long
    Dim SliceNo As Long
   
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            For SliceNo = 0 To (EndLevel - StartLevel) / Step
                OutVar(site, SliceNo) = Int((StartLevel + SliceNo * Step) / Lsb(site) / Kco) - 1
            Next SliceNo
        End If
    Next site

End Function

Public Function MakeSliceLevel_ShiroKobu_2Step(ByRef OutVar() As Double, ByVal StartLevel_Lo As Double, ByVal EndLevel_Lo As Double, ByVal Step_Lo As Double, ByVal StartLevel_Hi As Double, ByVal EndLevel_Hi As Double, ByVal Step_Hi As Double, Optional ByVal Lsb As Variant = 1, Optional ByVal Kco As Double = 1)

    Dim site As Long
    Dim SliceNo As Long

    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            For SliceNo = 0 To (EndLevel_Lo - StartLevel_Lo) / Step_Lo + (EndLevel_Hi - StartLevel_Hi) / Step_Hi
                If SliceNo <= (EndLevel_Lo - StartLevel_Lo) / Step_Lo Then
                    OutVar(site, SliceNo) = Int((StartLevel_Lo + SliceNo * Step_Lo) / Lsb(site) / Kco) - 1
                Else
                    OutVar(site, SliceNo) = Int((StartLevel_Hi + (SliceNo - (EndLevel_Lo - StartLevel_Lo) / Step_Lo) * Step_Hi) / Lsb(site) / Kco) - 1
                End If
            Next SliceNo
        End If
    Next site

End Function

Public Function Count_ShiroKobu(ByVal srcPlane As CImgPlane, ByVal srcZone As String, ByRef limArray() As Double, ByRef OutVar() As Double)

    Dim i As Long
    Dim site As Long
    Dim tmp_Slice(nSite) As Double
    Dim tmp_Count(nSite) As Double
    
    Dim tmp_Count_Array() As Double
    ReDim tmp_Count_Array(nSite, UBound(limArray, 2)) As Double

    For i = 0 To UBound(limArray, 2)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                tmp_Slice(site) = limArray(site, i)
            End If
        Next site

        Call Count(srcPlane, srcZone, EEE_COLOR_ALL, idpCountAbove, tmp_Slice, tmp_Slice, idpLimitExclude, tmp_Count)

        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                tmp_Count_Array(site, i) = tmp_Count(site)
            End If
        Next site
    Next i

    For i = 0 To UBound(limArray, 2)
        For site = 0 To nSite
            If i < UBound(limArray, 2) Then
                OutVar(site, i) = tmp_Count_Array(site, i) - tmp_Count_Array(site, i + 1)
            Else
                OutVar(site, i) = tmp_Count_Array(site, i)
            End If
        Next site
    Next i

End Function

Public Function ResultAdd_ShiroKobu(ByRef TsetLabel As String, ByRef Output() As Double, ByRef Num As Long)

    Dim site As Long
    Dim tmp1(nSite) As Double
    
    For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                    tmp1(site) = Output(site, Num)
            End If
    Next site

    If TheResult.IsExist(TsetLabel) = True Then
        TheResult.Delete (TsetLabel)
    End If
    
    Call ResultAdd(TsetLabel, tmp1)

End Function

'' 以下、13/10/29追加関数
Public Function WritePixelAddrSite(ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef pDataArray() As CPixInfo, _
                                    Optional Err_Code As Double = 999)

    Dim site As Long
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            If pDataArray(site).OverFlow_Flag Then
'                MsgBox "OverFlow PixData WritePix"
                Call dstPlane.SetPMD(dstPlane.BasePMD.Name)
                Call dstPlane.WritePixel(Err_Code, , site)
            ElseIf pDataArray(site).Exist_Flag Then
                '' Call dstPlane.SetPMD(dstZone)
                Call dstPlane.SetPMD(dstPlane.BasePMD.Name)
                Call dstPlane.WritePixelAddr(pDataArray(site).ALLPixInfo, site)
            End If
        End If
    Next site

End Function

Public Sub ReadPixelSite( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, _
    ByRef dataNum As Variant, ByVal pFlgName As String, _
    ByRef retPixArr() As CPixInfo, ByRef AddrMode As IdpAddrMode, _
    Optional ByVal MaxCount As Double = -1 _
)
    
        
    
    Dim site As Long
    Dim tmpRPD() As T_PIXINFO
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            Set retPixArr(site) = New CPixInfo
            If MaxCount < 0 Then MaxCount = retPixArr(site).CpixMaxCount
            If dataNum(site) > MaxCount Then
'                MsgBox "ReadPixData_MaxCount is :" & MaxCount
                retPixArr(site).OverFlow_Flag = True
            ElseIf dataNum(site) > 0 Then
                Call srcPlane.SetPMD(srcZone)
                Call srcPlane.PixelLog(site, pFlgName, tmpRPD, dataNum(site), AddrMode)
                Call retPixArr(site).SetPixInfo(tmpRPD)
            End If
        End If
    Next site
End Sub


Public Sub ReadPixelSite_FlagPlane( _
    ByRef srcPlaneGroup As String, ByVal srcZone As Variant, _
    ByVal pFlgName As String, _
    ByRef retPixArr() As CPixInfo, ByRef AddrMode As IdpAddrMode, _
    Optional ByVal MaxCount As Double = -1 _
)
    
    Dim site As Long
    Dim tmpRPD() As T_PIXINFO
    Dim tmpPlane As CImgPlane
    Set tmpPlane = TheIDP.PlaneManager(srcPlaneGroup).GetSharedFlagPlane(pFlgName).FlgPlane
    Dim dataNum(nSite) As Double
    
    Call tmpPlane.SetPMD(srcZone)
    Call tmpPlane.Num(dataNum, EEE_COLOR_ALL, pFlgName)
    
    Dim i As Long
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            Set retPixArr(site) = New CPixInfo
            If MaxCount < 0 Then MaxCount = retPixArr(site).CpixMaxCount
            If dataNum(site) > MaxCount Then
'                MsgBox "ReadPixData_MaxCount is :" & MaxCount
                retPixArr(site).OverFlow_Flag = True
            ElseIf dataNum(site) > 0 Then
                Call tmpPlane.PixelLog(site, pFlgName, tmpRPD, dataNum(site), AddrMode)
                For i = 0 To UBound(tmpRPD)
                    tmpRPD(0).Value = 1
                Next i
                Call retPixArr(site).SetPixInfo(tmpRPD)
            End If
        End If
    Next site
End Sub


Public Sub RPDOffset(result() As CPixInfo, inRPD() As CPixInfo, _
                        Optional offsetX As Double = 0, _
                        Optional offsetY As Double = 0, _
                        Optional gainX As Double = 1, _
                        Optional gainY As Double = 1, _
                        Optional putValue As Double = -1)

    Dim site As Long
    Dim i As Long
    Dim max As Long
    Dim tempPix() As T_PIXINFO
    
    For site = 0 To nSite
        Set result(site) = New CPixInfo
        If TheExec.sites.site(site).Active = True Then
            If inRPD(site).OverFlow_Flag Then
                result(site).OverFlow_Flag = True
            ElseIf inRPD(site).Exist_Flag Then
                ReDim tempPix(inRPD(site).Count - 1)
    
                For i = 0 To inRPD(site).Count - 1
                    tempPix(i).x = inRPD(site).PixInfo(i).x * gainX + offsetX
                    tempPix(i).y = inRPD(site).PixInfo(i).y * gainY + offsetY
                    If putValue = -1 Then
                        tempPix(i).Value = inRPD(site).PixInfo(i).Value
                    Else
                        tempPix(i).Value = putValue
                    End If
                Next i
                Call result(site).SetPixInfo(tempPix)
            End If
        End If
    Next site
End Sub

Public Sub RPDUnion(result() As CPixInfo, ParamArray var() As Variant)

    Dim site As Long
    Dim i As Long, j As Long
    Dim max As Long
    Dim my_OverFlowFlag As Boolean
    For site = 0 To nSite
        If result(site) Is Nothing Then Set result(site) = New CPixInfo
        If TheExec.sites.site(site).Active = True Then
            
            max = 0
            my_OverFlowFlag = False
            For i = 0 To UBound(var)
                If var(i)(site).Exist_Flag Then max = max + UBound(var(i)(site).ALLPixInfo) + 1
                my_OverFlowFlag = my_OverFlowFlag Or var(i)(site).OverFlow_Flag
            Next i
            
            If my_OverFlowFlag Then
                result(site).OverFlow_Flag = True
            ElseIf max <> 0 Then
                result(site).OverFlow_Flag = False
                max = max - 1
                
                Dim tempC As CPixInfo
                Dim Count As Long
                Count = 0
                Dim tempPix() As T_PIXINFO
                ReDim tempPix(max)
                For i = 0 To UBound(var)
                    Set tempC = var(i)(site)
                    If tempC.Exist_Flag Then
                        For j = 0 To UBound(tempC.ALLPixInfo)
                            tempPix(Count) = tempC.PixInfo(j)
                            Count = Count + 1
                        Next j
                    End If
                Next i
                Call result(site).SetPixInfo(tempPix)
            End If
        End If
    Next site
End Sub


Private Function GetColorList(ColorList() As String, ByVal dataArr As Variant) As String()

    If (dataArr(0) = "-") Then
        GetColorList = ColorList
    Else
        ReDim result(UBound(dataArr)) As String
        Dim i As Integer
        For i = 0 To UBound(dataArr)
            result(i) = dataArr(i)
        Next i
        GetColorList = result
    End If
End Function

Public Function mf_CountRandomVline( _
    ByVal deviationPlane As CImgPlane, _
    ByVal deviationZone As String, _
    ByVal sliceLevel As Double, _
    ByVal numberBitShift As Long, _
    ByRef lsbValue() As Double) As Double()

    Dim site As Long
    Dim returnResult(nSite) As Double
    Dim sliceLevels(nSite) As Double

    'スライスレベルを求めます。比較対照画像の次元が自乗されているので、こちらも自乗しておきます。
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            sliceLevels(site) = CLng((sliceLevel / lsbValue(site)) * 2 ^ numberBitShift - 0.5)
        End If
    Next site
    'カウントします。
    Call Count(deviationPlane, deviationZone, idpColorFlat, idpCountAbove, sliceLevels, sliceLevels, idpLimitExclude, returnResult)

    mf_CountRandomVline = returnResult

End Function

Public Sub OFD_Mask(ByRef srcPlane As CImgPlane, ByRef dstPlane As CImgPlane, ByVal pSubPmdName As String, ParamArray Zone() As Variant)

    Dim i As Long
    Dim pX As Long
    Dim pY As Long
    Dim pWidth As Long
    Dim pHeight As Long
    Dim X_Start As Long
    Dim X_End As Long
    Dim Y_Start As Long
    Dim Y_End As Long
    Dim tmp_X_Start As Long
    Dim tmp_X_End As Long
    Dim tmp_Y_Start As Long
    Dim tmp_Y_End As Long

    If UBound(Zone) = 0 Then
        Call Copy(srcPlane, srcPlane.BasePMD, EEE_COLOR_FLAT, dstPlane, srcPlane.BasePMD, EEE_COLOR_FLAT)
        Exit Sub
    End If

    tmp_X_Start = TheIDP.PMD(Zone(0)).Left
    tmp_X_End = TheIDP.PMD(Zone(0)).Right
    tmp_Y_Start = TheIDP.PMD(Zone(0)).Top
    tmp_Y_End = TheIDP.PMD(Zone(0)).Bottom

    For i = 1 To UBound(Zone)
        X_Start = TheIDP.PMD(Zone(i)).Left
        X_End = TheIDP.PMD(Zone(i)).Right
        Y_Start = TheIDP.PMD(Zone(i)).Top
        Y_End = TheIDP.PMD(Zone(i)).Bottom

        'X_Start = MIN
        If tmp_X_Start > X_Start Then tmp_X_Start = X_Start
        'X_End = MAX
        If tmp_X_End < X_End Then tmp_X_End = X_End
        'Y_Start = MIN
        If tmp_Y_Start > Y_Start Then tmp_Y_Start = Y_Start
        'Y_End = MAX
        If tmp_Y_End < Y_End Then tmp_Y_End = Y_End
    Next i

    'Create Zone
    pX = tmp_X_Start
    pY = tmp_Y_Start
    pWidth = tmp_X_End - tmp_X_Start + 1
    pHeight = tmp_Y_End - tmp_Y_Start + 1

    If TheIDP.isExistingPMD(pSubPmdName) = False Then
        Call TheIDP.PlaneManager(srcPlane.planeGroup).CreateSubPMD(pSubPmdName, pX, pY, pWidth, pHeight)
    End If


    Dim ofdPmdName As String
    ofdPmdName = "OFD"
    For i = 0 To UBound(Zone)
        ofdPmdName = ofdPmdName & "_" & Zone(i)
    Next i

    If TheIDP.isExistingPMD("LEFT_" & ofdPmdName) = False Then

        Call dstPlane.SetPMD(pSubPmdName).WritePixel(0)
        For i = 0 To UBound(Zone)
            Call dstPlane.SetPMD(Zone(i)).WritePixel(1)
        Next i

        Dim AreaSize_X As Long, AreaSize_Y As Long
        AreaSize_X = TheIDP.PMD(pSubPmdName).width
        AreaSize_Y = TheIDP.PMD(pSubPmdName).height
        Dim site As Integer
        Dim tmp() As Double
        For site = 0 To nSite
            Call dstPlane.SetPMD(pSubPmdName)
            Call dstPlane.ReadPixel(site, tmp, , (AreaSize_X * AreaSize_Y))
            Exit For
        Next site

        Dim Sum_X_Start As Long
        Dim Sum_X_End As Long
        Dim Sum_Y_Start As Long
        Dim Sum_Y_End As Long
        Sum_X_Start = TheIDP.PMD(pSubPmdName).Left
        Sum_X_End = TheIDP.PMD(pSubPmdName).Right
        Sum_Y_Start = TheIDP.PMD(pSubPmdName).Top
        Sum_Y_End = TheIDP.PMD(pSubPmdName).Bottom

        Dim pmdArray() As Integer
        ReDim pmdArray((AreaSize_X + Sum_X_Start), (AreaSize_Y + Sum_Y_Start))
        Dim x As Long, y As Long
        i = 0
        For y = Sum_Y_Start To (AreaSize_Y + Sum_Y_Start) - 1
            For x = Sum_X_Start To (AreaSize_X + Sum_X_Start) - 1
                pmdArray(x, y) = tmp(i)
                i = i + 1
            Next x
        Next y

        Dim Zone3Name As String
        Zone3Name = ""
        Dim zone3Find As Boolean
        For i = 0 To UBound(Zone)
            zone3Find = Zone(i) Like "*ZONE3"
            If zone3Find = True Then
                Zone3Name = Zone(i)
                Exit For
            End If
        Next i
        If Zone3Name = "" Then MsgBox ("ERROR!")

        Dim ZONE3_X_Start As Long
        Dim ZONE3_X_End As Long
        Dim ZONE3_Y_Start As Long
        Dim ZONE3_Y_End As Long
        ZONE3_X_Start = TheIDP.PMD(Zone3Name).Left
        ZONE3_X_End = TheIDP.PMD(Zone3Name).Right
        ZONE3_Y_Start = TheIDP.PMD(Zone3Name).Top
        ZONE3_Y_End = TheIDP.PMD(Zone3Name).Bottom

        Dim OFDL_Width As Long, OFDL_Height As Long
        OFDL_Width = 0
        x = ZONE3_X_Start - 1
        Do While x >= Sum_X_Start
            If pmdArray(x, ZONE3_Y_Start) = 1 Then
                Exit Do
            Else
                OFDL_Width = OFDL_Width + 1
                x = x - 1
            End If
        Loop
        OFDL_Height = TheIDP.PMD(Zone3Name).height

        Dim OFDR_Width As Long, OFDR_Height As Long
        OFDR_Width = 0
        x = ZONE3_X_End + 1
        Do While x <= Sum_X_End
            If pmdArray(x, ZONE3_Y_Start) = 1 Then
                Exit Do
            Else
                OFDR_Width = OFDR_Width + 1
                x = x + 1
            End If
        Loop
        OFDR_Height = TheIDP.PMD(Zone3Name).height

        Dim OFDT_Width As Long, OFDT_Height As Long, OFDT_X As Long
        OFDT_Width = 0: OFDT_Height = 0: OFDT_X = 0
        y = ZONE3_Y_Start - 1
        Do While y >= Sum_Y_Start
            If pmdArray(ZONE3_X_Start, y) = 1 Then
                Exit Do
            Else
                If OFDT_Width = 0 Then
                    For i = Sum_X_Start To Sum_X_End
                        If pmdArray(i, y) = 0 Then
                            If OFDT_X = 0 Then
                                OFDT_X = i
                            End If
                            OFDT_Width = OFDT_Width + 1
                        End If
                    Next i
                End If
                OFDT_Height = OFDT_Height + 1
                y = y - 1
            End If
        Loop

        Dim OFDB_Width As Long, OFDB_Height As Long, OFDB_X As Long
        OFDB_Width = 0: OFDB_Height = 0: OFDB_X = 0
        y = ZONE3_Y_End + 1
        Do While y <= Sum_Y_End
            If pmdArray(ZONE3_X_Start, y) = 1 Then
                Exit Do
            Else
                If OFDB_Width = 0 Then
                    For i = Sum_X_Start To Sum_X_End
                        If pmdArray(i, y) = 0 Then
                            If OFDB_X = 0 Then
                                OFDB_X = i
                            End If
                            OFDB_Width = OFDB_Width + 1
                        End If
                    Next i
                End If
                OFDB_Height = OFDB_Height + 1
                y = y + 1
            End If
        Loop


        Dim ColorMapWidth As Long
        Dim ColorMapHeight As Long
        ColorMapWidth = srcPlane.planeMap.width
        ColorMapHeight = srcPlane.planeMap.height

        'OFDL Create subPMD
        Dim OFDL_X As Long, OFDL_Y As Long, OFDL_offset As Long
        OFDL_X = ZONE3_X_Start - OFDL_Width
        OFDL_Y = ZONE3_Y_Start
        OFDL_offset = OFDL_Width Mod ColorMapWidth
        Call TheIDP.PlaneManager(srcPlane.planeGroup).CreateSubPMD("MASK_LEFT_" & ofdPmdName, OFDL_X, OFDL_Y, OFDL_Width, OFDL_Height)
        Call TheIDP.PlaneManager(srcPlane.planeGroup).CreateSubPMD("LEFT_" & ofdPmdName, (ZONE3_X_Start + OFDL_offset), OFDL_Y, OFDL_Width, OFDL_Height)

        'OFDR Create subPMD
        Dim OFDR_X As Long, OFDR_Y As Long, OFDR_offset As Long
        OFDR_X = ZONE3_X_End + 1
        OFDR_Y = ZONE3_Y_Start
        OFDR_offset = OFDR_Width Mod ColorMapWidth
        Call TheIDP.PlaneManager(srcPlane.planeGroup).CreateSubPMD("MASK_RIGHT_" & ofdPmdName, OFDR_X, OFDR_Y, OFDR_Width, OFDR_Height)
        Call TheIDP.PlaneManager(srcPlane.planeGroup).CreateSubPMD("RIGHT_" & ofdPmdName, (ZONE3_X_End - OFDR_Width - OFDR_offset + 1), OFDR_Y, OFDR_Width, OFDR_Height)

        'OFDT Create subPMD
        If OFDT_Height <> 0 Then
            Dim OFDT_Y As Long, OFDT_offset As Long
            OFDT_Y = ZONE3_Y_Start - OFDT_Height
            OFDT_offset = OFDT_Width Mod ColorMapHeight
            Call TheIDP.PlaneManager(srcPlane.planeGroup).CreateSubPMD("MASK_TOP_" & ofdPmdName, OFDT_X, OFDT_Y, OFDT_Width, OFDT_Height)
            Call TheIDP.PlaneManager(srcPlane.planeGroup).CreateSubPMD("TOP_" & ofdPmdName, OFDT_X, (ZONE3_Y_Start + OFDT_offset), OFDT_Width, OFDT_Height)
        End If

        'OFDB Create subPMD
        If OFDB_Height <> 0 Then
            Dim OFDB_Y As Long, OFDB_offset As Long
            OFDB_Y = ZONE3_Y_End + 1
            OFDB_offset = OFDB_Width Mod ColorMapHeight
            Call TheIDP.PlaneManager(srcPlane.planeGroup).CreateSubPMD("MASK_BOTTOM_" & ofdPmdName, OFDB_X, OFDB_Y, OFDB_Width, OFDB_Height)
            Call TheIDP.PlaneManager(srcPlane.planeGroup).CreateSubPMD("BOTTOM_" & ofdPmdName, OFDB_X, (ZONE3_Y_End - OFDB_Height - OFDB_offset + 1), OFDB_Width, OFDB_Height)
        End If

    End If

    'BaseCopy
    Call srcPlane.SetPMD(pSubPmdName)
    Call dstPlane.SetPMD(pSubPmdName)
    Call dstPlane.CopyPlane(srcPlane)

    'OFD LEFT Copy
    Call srcPlane.SetPMD("LEFT_" & ofdPmdName)
    Call dstPlane.SetPMD("MASK_LEFT_" & ofdPmdName)
    Call dstPlane.CopyPlane(srcPlane)

    'OFD RIGHT Copy
    Call srcPlane.SetPMD("RIGHT_" & ofdPmdName)
    Call dstPlane.SetPMD("MASK_RIGHT_" & ofdPmdName)
    Call dstPlane.CopyPlane(srcPlane)

    'OFD TOP Copy
    Dim tmpPlane As CImgPlane
    If TheIDP.isExistingPMD("TOP_" & ofdPmdName) = True Then
        Call GetFreePlane(tmpPlane, srcPlane.planeGroup, srcPlane.BitDepth, , "tmpPlane")
        Call tmpPlane.SetPMD(pSubPmdName)
        Call dstPlane.SetPMD(pSubPmdName)
        Call tmpPlane.CopyPlane(dstPlane)

        Call tmpPlane.SetPMD("TOP_" & ofdPmdName)
        Call dstPlane.SetPMD("MASK_TOP_" & ofdPmdName)
        Call dstPlane.CopyPlane(tmpPlane)
        Call ReleasePlane(tmpPlane)
    End If

    'OFD BOTTOM Copy
    If TheIDP.isExistingPMD("BOTTOM_" & ofdPmdName) = True Then
        Call GetFreePlane(tmpPlane, srcPlane.planeGroup, srcPlane.BitDepth, , "tmpPlane")
        Call tmpPlane.SetPMD(pSubPmdName)
        Call dstPlane.SetPMD(pSubPmdName)
        Call tmpPlane.CopyPlane(dstPlane)

        Call tmpPlane.SetPMD("BOTTOM_" & ofdPmdName)
        Call dstPlane.SetPMD("MASK_BOTTOM_" & ofdPmdName)
        Call dstPlane.CopyPlane(tmpPlane)
        Call ReleasePlane(tmpPlane)
    End If

End Sub

'' 0000131028 追記
Public Sub StdDev_FA( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)

    Call srcPlane.SetPMD(srcZone)
    
    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    If srcColor = EEE_COLOR_ALL Then
        Dim tempResult As ColorAllResult
        Call srcPlane.StdDevColorAll(tempResult, pFlgName)
    
        Call retResult.SetParam(srcPlane, tempResult)
        
    Else
        '' 現状、単色指定もFlatに代入されるが、色指定して代入することもおそらくできる。
        Dim temp() As Double
        Call srcPlane.StdDev(temp, srcColor, pFlgName)
        Call retResult.CreateFlat(temp)
    
    End If

End Sub

Public Sub Average_FA( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)

    Call srcPlane.SetPMD(srcZone)
    
    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
        
    If srcColor = EEE_COLOR_ALL Then
        Dim tempResut As ColorAllResult
        Call srcPlane.AverageColorAll(tempResut, pFlgName)

        Call retResult.SetParam(srcPlane, tempResut)
    
    Else
    
'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
        Dim temp() As Double
        Call srcPlane.Average(temp, srcColor, pFlgName)
        Call retResult.CreateFlat(temp)
    End If

End Sub

Public Sub sum_FA( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)

    Call srcPlane.SetPMD(srcZone)
    
    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    If srcColor = EEE_COLOR_ALL Then
        Dim tempResut As ColorAllResult
        Call srcPlane.SumColorAll(tempResut, pFlgName)
    
        Call retResult.SetParam(srcPlane, tempResut)

    Else

'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
        Dim temp() As Double
        Call srcPlane.sum(temp, srcColor, pFlgName)
        Call retResult.CreateFlat(temp)
    End If

End Sub

Public Sub GetPixelCount_FA( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)

    Call srcPlane.SetPMD(srcZone)
    
    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    
    If srcColor = EEE_COLOR_ALL Then
        Dim tempResut As ColorAllResult
        Call srcPlane.NumColorAll(tempResut, pFlgName)

        Call retResult.SetParam(srcPlane, tempResut)

    Else

'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
        Dim temp() As Double
        Call srcPlane.Num(temp, srcColor, pFlgName)
        Call retResult.CreateFlat(temp)
    End If

End Sub

Public Sub Min_FA( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
    
    Call srcPlane.SetPMD(srcZone)
    
    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    If srcColor = EEE_COLOR_ALL Then
        Dim tempResut As ColorAllResult
        Call srcPlane.MinColorAll(tempResut, pFlgName)
        
        Call retResult.SetParam(srcPlane, tempResut)

    Else

'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
        Dim temp() As Double
        Call srcPlane.Min(temp, srcColor, pFlgName)
        Call retResult.CreateFlat(temp)
    End If

End Sub

Public Sub Max_FA( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)
    
    Call srcPlane.SetPMD(srcZone)
    
    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    If srcColor = EEE_COLOR_ALL Then
        Dim tempResut As ColorAllResult
        Call srcPlane.MaxColorAll(tempResut, pFlgName)
        
        Call retResult.SetParam(srcPlane, tempResut)

    Else

'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
        Dim temp() As Double
        Call srcPlane.max(temp, srcColor, pFlgName)
        Call retResult.CreateFlat(temp)
    End If

End Sub
Public Sub MinMax_FA( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retMin As CImgColorAllResult, ByRef retMax As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)

    Call srcPlane.SetPMD(srcZone)

    If retMax Is Nothing Then
        Set retMax = New CImgColorAllResult
    End If

    If retMin Is Nothing Then
        Set retMin = New CImgColorAllResult
    End If


    If srcColor = EEE_COLOR_ALL Then
        Dim tempResutMin As ColorAllResult
        Dim tempResutMax As ColorAllResult
        Call srcPlane.MinMaxColorAll(tempResutMin, tempResutMax, pFlgName)
    
        Call retMax.SetParam(srcPlane, tempResutMax)
    
        Call retMin.SetParam(srcPlane, tempResutMin)

    Else

'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
'
        Dim temp_min() As Double
        Dim temp_max() As Double

        Call srcPlane.MinMax(temp_min, temp_max, srcColor, pFlgName)
        Call retMin.CreateFlat(temp_min)
        Call retMax.CreateFlat(temp_max)
    End If

End Sub

Public Sub DiffMinMax_FA( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)

    Call srcPlane.SetPMD(srcZone)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If

    If srcColor = EEE_COLOR_ALL Then
        Dim tempResut As ColorAllResult
        Call srcPlane.DiffMinMaxColorAll(tempResut, pFlgName)
        
        Call retResult.SetParam(srcPlane, tempResut)

    Else

'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
        Dim temp() As Double
        Call srcPlane.DiffMinMax(temp, srcColor, pFlgName)
        Call retResult.CreateFlat(temp)
    End If

End Sub

Public Sub AbsMax_FA( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByRef retResult As CImgColorAllResult, _
    Optional pFlgName As String = "" _
)

    Call srcPlane.SetPMD(srcZone)

    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If

    If srcColor = EEE_COLOR_ALL Then
        Dim tempResut As ColorAllResult
        Call srcPlane.AbsMaxColorAll(tempResut, pFlgName)
    
        Call retResult.SetParam(srcPlane, tempResut)
    
    Else

'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
        Dim temp() As Double
        Call srcPlane.AbsMax(temp, srcColor, pFlgName)
        Call retResult.CreateFlat(temp)
    End If

End Sub

Public Sub count_FA( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal countType As IdpCountType, _
    ByVal loLim As Variant, ByVal hiLim As Variant, ByVal pCountLimMode As IdpCountLimitMode, ByVal limitType As IdpLimitType, ByRef retResult As CImgColorAllResult, _
    Optional ByVal pFlgName As String = "", Optional ByVal pInputFlgName As String = "" _
)

    Call srcPlane.SetPMD(srcZone)
    
    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    If srcColor = EEE_COLOR_ALL Then
        Dim tempResut As ColorAllResult
        Call srcPlane.CountColorAll(tempResut, countType, loLim, hiLim, pCountLimMode, limitType, pFlgName, pInputFlgName)
    
        Call retResult.SetParam(srcPlane, tempResut)

    Else

'   hiLim,loLimをVariantで定義しているのは、サイト別で境界値が違う場合に対応するため。
'   サイト配列を入れると各サイトごと別々に、定数を入れると全サイトに同じ値が適用される。
'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
        Dim temp() As Double
        Call srcPlane.Count(temp, countType, loLim, hiLim, limitType, srcColor, pFlgName, pInputFlgName)
        Call retResult.CreateFlat(temp)
    End If

End Sub

Public Sub CountForFlgBitImgPlane_FA( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal countType As IdpCountType, _
    ByVal loLim As Variant, ByVal hiLim As Variant, ByVal pCountLimMode As IdpCountLimitMode, ByVal limitType As IdpLimitType, ByRef retResult As CImgColorAllResult, _
    ByRef pFlgPlane As CImgPlane, ByVal pFlgBit As Long, _
    Optional ByVal pInputFlgName As String = "" _
)

    Call srcPlane.SetPMD(srcZone)
    Call pFlgPlane.SetPMD(srcZone)
    
    If retResult Is Nothing Then
        Set retResult = New CImgColorAllResult
    End If
    
    If srcColor = EEE_COLOR_ALL Then
        Dim tempResut As ColorAllResult
        Call srcPlane.CountColorAllForFlgBitImgPlane(tempResut, countType, loLim, hiLim, pCountLimMode, pFlgPlane, pFlgBit, limitType, pInputFlgName)
        
        Call retResult.SetParam(srcPlane, tempResut)

    Else

'備考:
'   hiLim,loLimをVariantで定義しているのは、サイト別で境界値が違う場合に対応するため。
'   サイト配列を入れると各サイトごと別々に、定数を入れると全サイトに同じ値が適用される。
'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
        Dim temp() As Double
        Call srcPlane.CountForFlgBitImgPlane(temp, countType, loLim, hiLim, pFlgPlane, pFlgBit, limitType, srcColor, pInputFlgName)
        Call retResult.CreateFlat(temp)
    End If

End Sub

Public Sub PutFlag_FA( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, ByVal countType As IdpCountType, _
    ByVal loLim As Variant, ByVal hiLim As Variant, ByVal pCountLimMode As IdpCountLimitMode, ByVal limitType As IdpLimitType, _
    ByVal pFlgName As String, Optional ByVal pInputFlgName As String _
)
    Call srcPlane.SetPMD(srcZone)
    
    If srcColor = EEE_COLOR_ALL Then
        Call srcPlane.PutFlagColorAll(pFlgName, countType, loLim, hiLim, pCountLimMode, limitType, pInputFlgName)

    Else

'備考:
'   hiLim,loLimをVariantで定義しているのは、サイト別で境界値が違う場合に対応するため。
'   サイト配列を入れると各サイトごと別々に、定数を入れると全サイトに同じ値が適用される。
'   srcColorにEEE_COLOR_ALLを指定するとAllで処理をしてFlatの結果を返す。
'
        Call srcPlane.PutFlag(pFlgName, countType, loLim, hiLim, limitType, srcColor, pInputFlgName)

    End If

End Sub

Function CImgColorAllResultAcc(OutVar As CImgColorAllResult, inValue As CImgColorAllResult, ope As String, var As Variant)

    Dim site As Long
    Dim i As Long
    Dim varsType As String
    Dim tmp As Double
    Dim tmpVar As Double

    Set OutVar = inValue.Clone
    varsType = TypeName(var)
    
    If IsObject(var) Then
'        If Not StrComp(varsType, "CImgColorAllResult", vbTextCompare) Then MsgBox ("Object Cant Acc")
        Set OutVar = var.Clone
        varsType = "All"
    ElseIf IsArray(var) Then
        varsType = "Array"
    ElseIf IsNumeric(var) Then
        varsType = "Value"
    Else
        '' パラメータとして予期していない入力の場合
        MsgBox ("This param Cant Acc")
    End If
    
    
    Dim mf_colorList() As String
    mf_colorList = GetColorList(inValue.ColorList, Array("-"))

    '' 値が入ってきたときは、一度だけ代入
    If varsType = "Value" Then tmpVar = var
    For site = 0 To nSite
        '' 配列が入ってきたときは、Site毎に代入
        If varsType = "Array" Then tmpVar = var(site)
        For i = 0 To UBound(mf_colorList)
            '' CimgColorAllResult型が入ってきたときは、色毎に代入
            If varsType = "All" Then tmpVar = OutVar.color(mf_colorList(i)).SiteValue(site)
            '' 入力変数から
            tmp = inValue.color(mf_colorList(i)).SiteValue(site)
            '' 掛け算をする側の型を見る
            Select Case ope
                Case "+": tmp = tmp + tmpVar
                Case "-": tmp = tmp - tmpVar
                Case "*": tmp = tmp * tmpVar
                Case "/": tmp = Div(tmp, tmpVar, -999) '' エラーコードはこれでよい？
            End Select
            
            '' 出力変数にセット
            Call OutVar.SetData(mf_colorList(i), site, tmp)
        Next i
    Next site

    Set CImgColorAllResultAcc = OutVar

End Function

Public Function FlagColorExtract(refPlane As CImgPlane, flagName As String, ParamArray var() As Variant)

    If var(0) = "-" Then Exit Function
    
    Dim i As Long
    Call refPlane.GetSharedFlagPlane(flagName).SetFlagBit("FlagColorExtract_Temp")
    For i = 0 To UBound(var)
        Call SharedFlagOr(refPlane.planeGroup, refPlane.BasePMD, "FlagColorExtract_Temp", flagName, "FlagColorExtract_Temp", var(i))
    Next i
    Call SharedFlagAnd(refPlane.planeGroup, refPlane.BasePMD, flagName, flagName, "FlagColorExtract_Temp", EEE_COLOR_ALL)
    Call refPlane.GetSharedFlagPlane("FlagColorExtract_Temp").RemoveFlagBit("FlagColorExtract_Temp")

End Function

Public Sub Convolution_Clear( _
    ByRef srcPlane As CImgPlane, ByVal srcZone As Variant, ByRef srcColor As Variant, _
    ByRef dstPlane As CImgPlane, ByVal dstZone As Variant, ByRef dstColor As Variant, _
    ByVal Kernel As String, Optional ByVal divVal As Long = 0 _
)
    
    Dim tmpPlane As CImgPlane
    Call GetFreePlane(tmpPlane, srcPlane.planeGroup, srcPlane.BitDepth, True)
    Call tmpPlane.SetPMD(srcZone)
    Call srcPlane.SetPMD(srcZone)
    
    Call Copy(srcPlane, srcZone, EEE_COLOR_ALL, tmpPlane, srcZone, EEE_COLOR_ALL)
    
    Call dstPlane.SetPMD(dstZone)
    Call dstPlane.Convolution(tmpPlane, Kernel, dstColor, srcColor)

    If divVal <> 0 Then
        Call DivideConst(dstPlane, dstZone, dstColor, divVal, dstPlane, dstZone, dstColor)
    End If

End Sub

Public Function CleateDummyFlagPlane()

        
    Dim i As Integer
    Dim sPlane As CImgPlane
    Dim resPlane As CImgPlane
    Dim resBit As Integer
    Dim dstPlane As CImgPlane
    Dim dstBit As Integer
    
    Dim pFlgName As String
    pFlgName = "ForIMGDummyFlag"
    
    Dim pGroupName As String
        
    For i = 1 To TheIDP.PlaneManagerCount
        pGroupName = TheIDP.PlaneManager(i).Name
        If Left(pGroupName, 4) <> "CAP_" And Left(pGroupName, 1) <> "p" Then
            Call TheIDP.PlaneManager(i).GetSharedFlagPlane(pFlgName).SetFlagBit(pFlgName)
        End If
    Next i

End Function

Public Function ClearALLFlagBit(pFlgName As String)
    
    Dim i As Integer
    For i = 1 To TheIDP.PlaneManagerCount
        With TheIDP.PlaneManager(i)
            If .GetSharedFlagPlanes.Count > 0 Then
                If .GetSharedFlagPlane(1).Count > 0 Then
                    Call .GetSharedFlagPlane(pFlgName).RemoveFlagBit(pFlgName)
                End If
            End If
        End With
    Next i

End Function

Public Function Common_Log(ByVal inVar As Double, ByVal ErrorCode As Double) As Double

    If inVar <= 0 Then
        Common_Log = ErrorCode
    Else
        Common_Log = Log(inVar) / Log(10)
    End If

End Function

Public Sub SharedFlagOr_Array(ByRef pPlaneGroup As Variant, ByVal pZone As Variant, ByVal pDstName As String, _
        ParamArray pSrcName() As Variant)
    
    Dim i As Double
    
    With TheIDP.PlaneManager(Var2PlaneNameFlag(pPlaneGroup)).GetSharedFlagPlane(pDstName)
        Call .SetPMD(pZone)
        If UBound(pSrcName) = 0 Then
            Call .LOr(pDstName, pSrcName(0), pSrcName(0), EEE_COLOR_ALL)
        Else
            Call .LOr(pDstName, pSrcName(0), pSrcName(1), EEE_COLOR_ALL)
            For i = 2 To UBound(pSrcName)
                Call .LOr(pDstName, pDstName, pSrcName(i), EEE_COLOR_ALL)
            Next i
        End If
    End With

End Sub

Private Function Var2PlaneNameFlag(ByVal pVal As Variant) As String
    
    If IsObject(pVal) Then
        Var2PlaneNameFlag = pVal.Manager.Name
    Else
        Var2PlaneNameFlag = pVal
    End If
End Function

Public Function d_read_vmcu_Kernel(ByRef srcPlane As CImgPlane, _
                                  ByVal Zone As String, _
                                  ByVal MaxCount As Double, _
                                  ByVal signature As String, _
                                  ByVal Unit As String, _
                                  ByVal countType As String, _
                                  ByVal PointLevel As Variant, _
                                  ByVal CountLimit As String, _
                                  ByVal pColor As String, _
                                  Optional InputFlg As String = "", _
                                  Optional kernelName As String = "")

'' 連続点に色情報を持たせるのはハードルが高い。
    Dim site As Long
    Dim tmp_count_All As CImgColorAllResult
    Dim tmp_Count(nSite) As Double
    Dim tmp_count_color(nSite) As Double
    Dim ColorMap() As String
    Dim PixelLogResult() As T_PIXINFO
    Dim x As Long, y As Long, Data As Double
    Dim currentSiteDeviceNumber As String
    Dim i As Long
    Dim j As Long

    Call CountColorAll(srcPlane, Zone, countType, PointLevel, PointLevel, idpLimitEachSite, CountLimit, tmp_count_All, "PointDef", InputFlg)
    
    '' 本来、srcPlaneにはコンボリューションをかける前の情報を入力し、
    '' InputFlagにコンボリューション後のカウント結果を入れるようにし、
    '' PointLevelに1を指定すれば、動作するはず。
    '' 一旦、そのようなフローとなっていないため、kernelname を強制的に""とする。
    kernelName = ""
    If kernelName <> "" Then
        Set tmp_count_All = Nothing
        Call ReCreateFlag(srcPlane, Zone, pColor, kernelName, "PointDef", , PointLevel)
        Call CountColorAll(srcPlane, Zone, countType, 1, 1, idpLimitEachSite, idpLimitInclude, tmp_count_All, , "PointDef")
    End If
    
    Call GetSum_Color_Flat(tmp_Count, tmp_count_All)
    
    For site = 0 To nSite
        If tmp_Count(site) < MaxCount Then
            tmp_Count(site) = tmp_Count(site)
        Else
            tmp_Count(site) = MaxCount
        End If
    Next site
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
        
            If tmp_Count(site) <= 0 Then GoTo Next_Site
            currentSiteDeviceNumber = getActiveSiteDeviceNumer(site)
            
            ReDim PixelLogResult(tmp_Count(site))
            
            With srcPlane
                Call .SetPMD(Zone)
                Call .PixelLog(site, "PointDef", PixelLogResult, tmp_Count(site), idpAddrAbsolute)
            End With

            If Sw_Ana = 1 Or Flg_Debug = 1 Then
                TheExec.Datalog.WriteComment "*****  DEFECT " & signature & " DATA (SITE:" & site & ") *****"
            End If

            For i = 0 To tmp_Count(site) - 1
                x = PixelLogResult(i).x
                y = PixelLogResult(i).y
                Data = 1
                     
                If Sw_Ana = 1 Then
                    Print #m_fileLunDefectFile, returnCurrentDefectLine(currentSiteDeviceNumber, signature, Unit, x, y, Data, True)
                End If
                If Flg_Debug = 1 Then
                    TheExec.Datalog.WriteComment returnCurrentDefectLine(currentSiteDeviceNumber, signature, Unit, x, y, Data, False)
                End If
            Next i
Next_Site:
        End If
    Next site
            
    Call srcPlane.GetSharedFlagPlane("PointDef").RemoveFlagBit("PointDef")
                       

End Function

Public Function ReCreateFlag(ByRef srcPlane As CImgPlane, ByRef srcZone As Variant, ByRef srcColor As Variant, _
    ByVal Kernel As String, ByVal srcFlag As String, Optional dstFlag As String = "", Optional sliceCount As Variant = 0)
    
    '' srcPlane => Convolution前のプレーン情報
    
    '' 結果フラグが記述されていなければ入力に上書き
    If dstFlag = "" Then dstFlag = srcFlag
    
'' フラグを復元しなくてもよい条件
'' 1:左右対称のフラグであること　⇒　カーネルを逆にしてみて文字列一致で確認
'' 2:2点連続であること　⇒　中央値と周辺の最小値(0を除く)の和がスライスCount以上であること
    
'' カーネル情報を反転させてConvolutionをかけなおすことで、同様の処理が可能。
'' その場合にも、Skip処理を入れる事で、処理を短縮することが可能。
'' 処理内容： CreateKernel(ExistCheck) => FlagCopy => Convolution => PutFlag => PutFlag(PutFlagを入力マスクに指定)
    Dim Check1 As Boolean, Check2 As Boolean
    Check1 = True: Check2 = True
    
    '' カーネル情報の取得
    Dim kerInfo As CImgKernelInfo
    Set kerInfo = TheIDP.KernelManager.Kernel(Kernel)
    
    Dim ksizeX As Double, ksizeY As Double
    Dim msizeX As Double, msizeY As Double
    
    ksizeX = (kerInfo.width - 1) / 2
    ksizeY = (kerInfo.height - 1) / 2
    
    '' カーネルが奇数でない場合は、この処理は不可能。
    If ksizeX <> Int(ksizeX) Or ksizeY <> Int(ksizeY) Then MsgBox "KernelMap not oddValue"
    
    If srcColor = EEE_COLOR_FLAT Then
        msizeX = 1
        msizeY = 1
    Else
        msizeX = srcPlane.planeMap.width
        msizeY = srcPlane.planeMap.height
    End If
    
    '' カーネルの情報を文字列に置き換え
    Dim valueList() As String
    valueList = Split(kerInfo.Value, " ")

    '' SliceがDoubleでくる可能性とSite配列の可能性があるので置き換え
    Dim sliceCountD As Double
    If IsArray(sliceCount) Then
        sliceCountD = sliceCount(0)
    Else
        sliceCountD = sliceCount
    End If

    '' 条件チェック
    Dim i As Double
    Dim maxi As Double
    maxi = Int(UBound(valueList) / 2)
    For i = 0 To maxi
        If valueList(i) <> valueList(UBound(valueList) - i) Then Check1 = False
        If valueList(i) <> 0 And valueList(i) + valueList(maxi + 1) Then Check2 = False
    Next i
        
    '' 条件をともに満たしている場合以下の処理をして関数を抜ける
    If Check1 And Check2 Then
        If dstFlag <> srcFlag Then
            Call SharedFlagOr(srcPlane.planeGroup, srcZone, dstFlag, srcFlag, srcFlag, EEE_COLOR_ALL)
        End If
        Exit Function
    End If
    
    '' 条件を満たしていない場合は、以下の処理が走る
    
    Dim tmpPix1(nSite) As CPixInfo
    Dim tmpPix2(nSite) As CPixInfo
    
    Call ReadPixelSite_FlagPlane(srcPlane.planeGroup, srcZone, srcFlag, tmpPix1, idpAddrAbsolute)
    
    Dim resultPix(nSite) As CPixInfo
        
    Dim loopCount As Double
    loopCount = UBound(valueList)
    If loopCount <> UBound(valueList) Then MsgBox "Err"
    Dim loopX As Double
    Dim loopY As Double
    For loopY = -ksizeY To ksizeY
        For loopX = -ksizeX To ksizeX
            If valueList(loopCount) <> 0 Then
                Call RPDOffset(tmpPix2, tmpPix1, loopX * msizeX, loopY * msizeY, , , 1)
                Call RPDUnion(resultPix, resultPix, tmpPix2)
            End If
            loopCount = loopCount - 1
        Next loopX
    Next loopY

    Dim tmpPlane As CImgPlane
    Call GetFreePlane(tmpPlane, srcPlane.planeGroup, idpDepthS16, True)
    Call WritePixelAddrSite(tmpPlane, srcPlane.BasePMD, resultPix)
    
    Call PutFlag(tmpPlane, srcZone, EEE_COLOR_ALL, idpCountAbove, 1, 1, idpLimitInclude, "FLAG_ReCreateFlag")
    Call PutFlag(srcPlane, srcZone, EEE_COLOR_ALL, idpCountAbove, 1, 1, idpLimitInclude, dstFlag, "FLAG_ReCreateFlag")
    Call ClearALLFlagBit("FLAG_ReCreateFlag")

End Function

Public Function RPDextraction(ByRef pa() As CPixInfo, ByRef pb() As CPixInfo, _
                                ByVal Direcrion As String, ByVal Start_Add As Long, _
                                ByVal Cycle As Long, ByVal times As Long, _
                                ByVal Num As Long, _
                                Optional Value As Double = 1, _
                                Optional elseValue As Double = 0)
    Dim site As Long
    Dim tpixInfo() As T_PIXINFO
    Dim i As Long
    Dim j As Long
    If Direcrion <> "H" And Direcrion <> "V" Then MsgBox "ERR Direction"
    Start_Add = Start_Add
    For site = 0 To nSite
        Set pb(site) = New CPixInfo
        
        If pa(site).OverFlow_Flag Then
            pb(site).OverFlow_Flag = True
        ElseIf pa(site).Exist_Flag Then
            pb(site).OverFlow_Flag = False
            ReDim tpixInfo(pa(site).Count) As T_PIXINFO
            j = 0
            '' tpixInfo = pa(site).ALLPixInfo
            For i = 0 To pa(site).Count - 1
                If Direcrion = "H" Then
                    If (Int((pa(site).PixInfo(i).x - Start_Add) / Cycle) Mod times) + 1 = Num Then
                        tpixInfo(j) = pa(site).PixInfo(i)
                        tpixInfo(j).Value = Value
                        pb(site).Exist_Flag = True
                        j = j + 1
                    Else
'                        tpixInfo(j).Value = elseValue
                    End If
                End If
                If Direcrion = "V" Then
                    If (Int((pa(site).PixInfo(i).y - Start_Add) / Cycle) Mod times) + 1 = Num Then
                        tpixInfo(j) = pa(site).PixInfo(i)
                        tpixInfo(j).Value = Value
                        pb(site).Exist_Flag = True
                        j = j + 1
                    Else
'                        tpixInfo(i).Value = elseValue
                    End If
                End If
            Next i
            If pb(site).Exist_Flag Then
                Call pb(site).SetPixInfo(tpixInfo)
            End If
        End If
    Next site

End Function

Public Function RPD_areaCount( _
    ByRef pa() As CPixInfo, _
    ByRef pb() As CPixInfo, _
    ByRef OutVar() As Double, _
    kernelName As String, _
    Optional ngc As Double = 0, _
    Optional revisingFlag As Boolean = False) As Double

'' カーネルで定義されたエリア(中心点)に対して、
'' ngcで定義された個数が存在している場合、カウントを行う処理。
'' 補正をかけながらカウントを行うモード(revisingFlag = True)と
'' 補正をかけずにカウントを行うモード(revisingFlag = False)が存在する。
'' 中心点の値に限らず、個数指定はngcで指定するものとし、
'' ngc=0の場合、中心点から見て、周辺にフラグが立っていれば、自身を含む
'' すべての点を対象とする。1以上はその個数を残してカウントする。
'' pb() は補正残し分のリードPixデータなので注意。(カウント対象ではない。)

    Dim kw As Double, kh As Double, kvRow As String, kv As Variant
    
    kw = TheIDP.Kernel(kernelName).width
    kh = TheIDP.Kernel(kernelName).height
    kvRow = TheIDP.Kernel(kernelName).Value
    kv = Split(kvRow, " ")

    Dim i As Long
    Dim site As Long

    '// カーネルから 0 以外の箇所をカウントする
    Dim kCount As Integer
    kCount = 0
    For i = 0 To UBound(kv)
        If (kv(i) <> 0) Then kCount = kCount + 1
    Next i

    Dim result As Double
    '// 基準点からの X/Y オフセットを配列で宣言する
    Dim xOffs() As Double, yOffs() As Double
    ReDim xOffs(kCount)
    ReDim yOffs(kCount)

    Dim kIndex As Double, x As Double, y As Double
    
    kIndex = 0
    For y = 0 To kh - 1
        For x = 0 To kw - 1
            If (kv(x + y * kw) <> 0) Then
                xOffs(kIndex) = x - (kw - 1) / 2
                yOffs(kIndex) = y - (kh - 1) / 2
                kIndex = kIndex + 1
            End If
        Next x
    Next y

    '// pa を順番に見ていく
    Dim pX As Double, pY As Double
    Dim px1 As Double, py1 As Double
    Dim pValue As Double, pValue1 As Double
    
    Dim nCount As Double
    Dim j As Double, k As Double
    
    Dim tpixInfo() As T_PIXINFO
    Dim tpixInfoRes() As T_PIXINFO
    Dim tpixInfo_B() As Boolean
    
    For site = 0 To nSite
        If pb(site) Is Nothing Then Set pb(site) = New CPixInfo
        result = 0
        tpixInfo = pa(site).ALLPixInfo
        ReDim tpixInfoRes(pa(site).Count) As T_PIXINFO
        ReDim tpixInfo_B(pa(site).Count) As Boolean
        
        If pa(site).OverFlow_Flag Then
'            MsgBox "ReadPixSite overFlow"
            pb(site).OverFlow_Flag = True
            OutVar(site) = 9999
        Else
            pb(site).OverFlow_Flag = False
            For i = 0 To pa(site).Count - 1
                '// 点欠陥のアドレス
                pX = tpixInfo(i).x
                pY = tpixInfo(i).y
                pValue = tpixInfo(i).Value
        
                If (pValue = 0) Then
                    '// 既に補正されている場合は無視
                Else
                    '// 隣接カウント
                    nCount = 0
    '                For j = i + 1 To UBound(tpixInfo)  '// 自分より後の点(自分より右下の点）を見ていく
                    For j = 0 To pa(site).Count - 1 ' 基準より前の点も見に行く
                        px1 = tpixInfo(j).x
                        py1 = tpixInfo(j).y
                        pValue1 = tpixInfo(j).Value
                        
                        If (pValue1 = 0) Or (px1 = pX And py1 = pY) Then
                            ' Skipされる条件()
                        Else
        
                            '// y,x 座標がカーネルサイズから外れている場合は、このj 以降は全てはずれているのでループ終了
                            If (pY - (kh - 1) / 2 <= py1 And py1 <= pY + (kh - 1) / 2) And _
                               (pX - (kw - 1) / 2 <= px1 And px1 <= pX + (kw - 1) / 2) Then
        
                                ' 隣接対象でaはない。
                                '// カーネル内の 0 以外の箇所に該当するかどうかチェック
                                For k = 0 To kCount
                                    If (pX + xOffs(k) = px1 And pY + yOffs(k) = py1) Then
                                        '該当する
                                        nCount = nCount + 1
                                        If (nCount >= ngc) Then
                                            'リミットを超えている場合(ダブルカウント防止)
                                            If tpixInfo_B(j) = False Then
                                                tpixInfoRes(result) = tpixInfo(j)
                                                result = result + 1
                                                tpixInfo_B(j) = True
                                            End If
                                            ' 補正しながらの場合は、0を代入し、次回以降カウント対象外
                                            If revisingFlag Then tpixInfo(j).Value = 0
        '                                        Debug.Print "x" & px1 & "y" & py1 & "result" & result
                                            
                                            'エリア内検索が１以下の時は、基準点もカウント対象
                                            If ngc <= 0 And tpixInfo(i).Value <> 0 Then
                                                If tpixInfo_B(i) = False Then
                                                    tpixInfoRes(result) = tpixInfo(i)
                                                    result = result + 1
                                                    tpixInfo_B(i) = True
                                                End If
                                                If revisingFlag Then tpixInfo(i).Value = 0
                                            End If
                                        Else
                                             tpixInfo(j).Value = 1
                                        End If
                                        ' 対象の点と一対一でマッチした場合はExitForする。
                                        Exit For
                                    End If
                                Next k
                            ' RPD復元後、データが整列されていない可能性があるので、すべてのPixデータを確認する。
                            ' 整列されている場合は、以下のExitForを実行すれば、若干処理が早くなる。
                            ' ElseIf py1 > py + (kh - 1) / 2 And px1 > px + (kw - 1) / 2 Then
                                ' Exit For
                            End If
                        End If
                    Next j
                End If
            Next i
        
            If result > 0 Then
                pb(site).Exist_Flag = True
                Call pb(site).SetPixInfo(tpixInfoRes)
            End If
            OutVar(site) = result
        End If
    Next site
    

End Function

Public Function Count_ShiroKobu_marge(ByVal srcPlane As CImgPlane, ByVal srcZone As String, ByRef limArray() As Double, ByRef OutVar() As Double, ByVal CountType_start As String, ByVal CountType_end As String)

    Dim i As Long
    Dim site As Long
    Dim tmp_Slice(nSite) As Double
    Dim tmp_Count(nSite) As Double
    
    Dim tmp_Count_Array() As Double
    ReDim tmp_Count_Array(nSite, UBound(limArray, 2)) As Double

    For i = 0 To UBound(limArray, 2)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                tmp_Slice(site) = limArray(site, i)
            End If
        Next site

        Call Count(srcPlane, srcZone, EEE_COLOR_ALL, idpCountAbove, tmp_Slice, tmp_Slice, idpLimitExclude, tmp_Count)

        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                tmp_Count_Array(site, i) = tmp_Count(site)
            End If
        Next site
    Next i

    If CountType_start = "Above" Then
        For i = 0 To UBound(limArray, 2)
            For site = 0 To nSite
                    OutVar(site, i) = tmp_Count_Array(site, i)
            Next site
        Next i
    End If
    
    If CountType_start = "Between" Then
        If CountType_end = "Above" Then
            For i = 0 To UBound(limArray, 2)
                For site = 0 To nSite
                    If i < UBound(limArray, 2) Then
                        OutVar(site, i) = tmp_Count_Array(site, i) - tmp_Count_Array(site, i + 1)
                    Else
                        OutVar(site, i) = tmp_Count_Array(site, i)
                    End If
                Next site
            Next i
        Else
            For i = 0 To UBound(limArray, 2) - 1
                For site = 0 To nSite
                        OutVar(site, i) = tmp_Count_Array(site, i) - tmp_Count_Array(site, i + 1)
                Next site
            Next i
        End If
    End If

End Function

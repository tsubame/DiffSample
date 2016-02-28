Attribute VB_Name = "Image_MasterFunctions_MURA"
Option Explicit

'クラシックムラ=============================================
'Hue補正係数
'色再現測定を行い、それから得た値を設定。
Private RyGain As Double
Private ByGain As Double
Private RyHue As Double
Private ByHue As Double

Private MagicNumber As Double
Private IsClsMuraFlatFieldingOn As Boolean

'---- Look Up Tables
Private LookUpTableArctan As String
'---- Convolution Kernels
Private LowPassFilterNameH As String
Private LowPassFilterNameV As String
'---- Path to FF
Private FlatFieldPath As String
'---- Kernel Taps
Private LpfKernel As String

'FFの各色平均値
Private refFlatFieldMeanR(nSite) As Double
Private RefFlatFieldMeanG(nSite) As Double
Private refFlatFieldMeanB(nSite) As Double


'Labムラ==================================================
Const PLANEBANK_FLATFIELD_R As String = "Flat Field R"
Const PLANEBANK_FLATFIELD_G As String = "Flat Field G"
Const PLANEBANK_FLATFIELD_B As String = "Flat Field B"

'---- Path to FF
Private LabFlatFieldPath As String

'---- Kernel Taps
Private LabKernel_LowPassH As String
Private LabKernel_LowPassV As String

'---- Lab Parameters
Public LDENL_FLAG_MEDIAN_TAPS As Long
Public LOCAL_DIFF_SIZE As Long
Public COMP_PRIMARY_TO_LOCAL As Long
Public LABZ2D_X_SIZE As Long
Public LABZ2D_Y_SIZE As Long
Public LABZ2D_AREA As Long


'B掃きムラ==================================================
Private Nbtd_FlatFieldPath As String

Private BHDifSize As Long
Private BHShadingFilterTap As Long

Public COMP_PRIMARY_TO_BHLOCAL As Long

'===========================================================
Private Y_Mean(nSite) As Double

'微分時の推奨微分距離
Private MINIMUM_DIF_SIZE_REQUIRED As Long
'GLOBALムラのLOCALムラに対する大きさ(比)
Private SIZE_RATIO_GLOBAL_OVER_LOCAL As Long

'******** 基本ユニットサイズ(1MHz相当)設定 *****************
Private UnitSize As Long

'******** YLINEパラメータ設定 *****************************
Private COMP_RGB_TO_Y As Long
Private yLineDif As Long
Private yLineCoef As Double
Private yLineCompPix As Long
Private yLineCount As Long

'******** YLOCAL/FRAMパラメータ設定 ***********************
Private COMP_YLINE_TO_YLOCAL As Long
Private yLocalDif As Long
Private yLocalCoef As Double
Private yLocalCount As Long

'******** YGLOBALパラメータ設定 ***************************
Private COMP_YLOCAL_TO_YGLOBAL As Long
Private yGlobDif As Long
Private yGlobCoef As Double
Private yGlobCount As Long

'******** YGLOBAL2パラメータ設定 ***************************
Private yGlob2HDif As Double
Private yGlob2VDif As Double
Private yGlob2Count As Long

'******** CLOCAL/FRAMパラメータ設定 ***********************
Private COMP_YLOCAL_TO_CLOCAL As Long
Private cLocalDif As Long
Private cLocalCoef As Double

'******** CGLOBALパラメータ設定 ***************************
Private COMP_CLOCAL_TO_CGLOBAL As Long
Private cGlobDif As Long
Private cGlobCoef As Double

'******** ゾーンマップパラメータ設定 ****************************
Private clampOpbZone As String
Private colorMapName As String

Private Kernel_LowPassH As String
Private Kernel_LowPassV As String
Private Kernel_YLine As String
Private Kernel_Ylocal As String
Private Kernel_YFrame2D As String
Private Kernel_YFrame2 As String
Private Kernel_YGlobal As String
Private lookUpTable_lut15 As String

Public Enum CMReturnType
  rNUM = 1
  rMAX = 2 '配列の要素数としても使用しているので、MAXを最大とすること
End Enum

Public Enum rgbColorArray
  red = 0
  green = 1
  blue = 2
End Enum

'===========================================================

Public Sub labProc_Initialize(ByVal pType As String, ByVal ZONE_FULL As Variant, ByVal ZONE_ZONE3 As Variant, ByVal clampZone As Variant)

    Dim site As Long
    Dim Flg_Active(nSite) As Long

    If TheIDP.KernelManager.IsExist("kernel_LowPassH_ColorFloat") = True Then Exit Sub  'CHECK
    
    'ALL SITE ACTIVE
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Flg_Active(site) = 1
        End If
    Next site
    TheExec.sites.SetAllActive (True)

    'To read parameters
    Call StdLab_GetMuraParameters

    Call ker_labProcKernel
    
    Dim flatPlane As CImgPlane
    Call GetFreePlane(flatPlane, pType, idpDepthS16, True, "Flat Field (temporary)")
    Call labProc_LoadFlatFieldData(flatPlane, ZONE_FULL)
    Call labProc_BayerSeparationFF(flatPlane, ZONE_ZONE3, clampZone)
    Call ReleasePlane(flatPlane)

    'SITE ACTIVE RETURN
    For site = 0 To nSite
        If Flg_Active(site) = 0 Then
            TheExec.sites.site(site).Active = False
        End If
    Next site

End Sub

Private Sub ker_labProcKernel()
    
    With TheIDP
        .CreateKernel "kernel_LowPassH_ColorFloat", idpKernelFloat, 15, 1, 0, LabKernel_LowPassH
        .CreateKernel "kernel_LowPassV_ColorFloat", idpKernelFloat, 1, 15, 0, LabKernel_LowPassV
    End With
    
End Sub


Public Sub labProc_LoadFlatFieldData( _
    ByRef flatPlane As CImgPlane, _
    ByVal flatZone As String)
    
    '[Edit Here] Filename Prefix Setting
    Const FF_FILE_PREFIX As String = "HL-"
    Dim site As Long
    
    Dim myTesterName As String
    myTesterName = Format(Sw_Node, "000")
    
    Dim myTesterSite As String
    Dim myFileName As String

    For site = 0 To nSite
        myTesterSite = "-Site-" & Format(site, "0")
        myFileName = FF_FILE_PREFIX & myTesterName & myTesterSite & ".stb"
        Call InPutImage(site, flatPlane, flatZone, LabFlatFieldPath & myTesterName & "\" & myFileName)
    Next site
    
End Sub

Private Sub labProc_BayerSeparationFF( _
    ByRef flatPlane As CImgPlane, ByVal ZONE_ZONE3 As Variant, ByVal clampZone As Variant)
    
    'Image planes
    Const MYPLANE_LOCAL As String = "pLabL"
    Const MYPLANE_BAYER As String = "pBayer"
    Const MYPLANE_BAYER_R As String = "pBayerR"
    Const MYPLANE_BAYER_G As String = "pBayerG"
    Const MYPLANE_BAYER_B As String = "pBayerB"
    Const MYPLANE_COMP_BAYER_R As String = MYPLANE_LOCAL
    Const MYPLANE_COMP_BAYER_G As String = "pcBayerG"
    Const MYPLANE_COMP_BAYER_B As String = "pcBayerB"
    
    Dim compFactor As Long
    compFactor = labProc_ReturnCompFactor
    
    'Local variables
    Dim Bclamp(nSite) As Double
    
    'OPB Clamp
    Dim clampPlane As CImgPlane
    Call GetFreePlane(clampPlane, flatPlane.planeGroup, idpDepthS16, False, "MURA/Clamp Image")
    Call Average(flatPlane, clampZone, EEE_COLOR_FLAT, Bclamp)
    Call SubtractConst(flatPlane, ZONE_ZONE3, EEE_COLOR_FLAT, Bclamp, clampPlane, ZONE_ZONE3, EEE_COLOR_FLAT)
    
    'RGB separation
    Dim bayerPlane As CImgPlane
    Call GetFreePlane(bayerPlane, MYPLANE_BAYER, idpDepthS16, False, "Bayer Plane")
    Call Copy(clampPlane, ZONE_ZONE3, EEE_COLOR_FLAT, bayerPlane, "BAYER_FULL", EEE_COLOR_FLAT)
    Call ReleasePlane(clampPlane)
    
    Dim bayerPlaneR As CImgPlane
    Call GetFreePlane(bayerPlaneR, MYPLANE_BAYER_R, idpDepthS16, False, "R extracted plane")
    Call Copy(bayerPlane, "BAYER_FULL", "R", bayerPlaneR, "BAYER_R_FULL", "R")
    
    Dim bayerPlaneG As CImgPlane
    Call GetFreePlane(bayerPlaneG, MYPLANE_BAYER_G, idpDepthS16, False, "G extracted plane")
    Call MultiMean(bayerPlane, "BAYER_FULL", "GR", bayerPlaneG, "BAYER_G_FULL", "GR", idpMultiMeanFuncMean, 1, 2)
    
    Dim bayerPlaneB As CImgPlane
    Call GetFreePlane(bayerPlaneB, MYPLANE_BAYER_B, idpDepthS16, False, "B extracted plane")
    Call Copy(bayerPlane, "BAYER_FULL", "B", bayerPlaneB, "BAYER_B_FULL", "B")
    Call ReleasePlane(bayerPlane)
    
    'Random noise reduction (median filter)
    Dim bayerPlaneWorkR As CImgPlane
    Call GetFreePlane(bayerPlaneWorkR, bayerPlaneR.planeGroup, idpDepthS16, False, "Work for R-extracted")
    Call Median(bayerPlaneR, "BAYER_R_FULL", EEE_COLOR_FLAT, bayerPlaneWorkR, "BAYER_R_FULL", EEE_COLOR_FLAT, 5, 1)
    Call Median(bayerPlaneWorkR, "BAYER_R_FULL", EEE_COLOR_FLAT, bayerPlaneR, "BAYER_R_FULL", EEE_COLOR_FLAT, 1, 5)
    Call ReleasePlane(bayerPlaneWorkR)
    
    Dim bayerPlaneWorkG As CImgPlane
    Call GetFreePlane(bayerPlaneWorkG, bayerPlaneG.planeGroup, idpDepthS16, False, "Work for G-extracted")
    Call Median(bayerPlaneG, "BAYER_G_FULL", EEE_COLOR_FLAT, bayerPlaneWorkG, "BAYER_G_FULL", EEE_COLOR_FLAT, 5, 1)
    Call Median(bayerPlaneWorkG, "BAYER_G_FULL", EEE_COLOR_FLAT, bayerPlaneG, "BAYER_G_FULL", EEE_COLOR_FLAT, 1, 5)
    Call ReleasePlane(bayerPlaneWorkG)
    
    Dim bayerPlaneWorkB As CImgPlane
    Call GetFreePlane(bayerPlaneWorkB, bayerPlaneB.planeGroup, idpDepthS16, False, "Work for B-extracted")
    Call Median(bayerPlaneB, "BAYER_B_FULL", EEE_COLOR_FLAT, bayerPlaneWorkB, "BAYER_B_FULL", EEE_COLOR_FLAT, 5, 1)
    Call Median(bayerPlaneWorkB, "BAYER_B_FULL", EEE_COLOR_FLAT, bayerPlaneB, "BAYER_B_FULL", EEE_COLOR_FLAT, 1, 5)
    Call ReleasePlane(bayerPlaneWorkB)

    'Multi-Mean for local
    Dim rISrcPlane As CImgPlane
    Dim gISrcPlane As CImgPlane
    Dim bISrcPlane As CImgPlane
    Dim gCompPlane As CImgPlane
    Dim bCompPlane As CImgPlane
    Call GetFreePlane(rISrcPlane, MYPLANE_LOCAL, idpDepthS16, False, "R (int) Source Plane")
    Call GetFreePlane(gISrcPlane, MYPLANE_LOCAL, idpDepthS16, False, "G (int) Source Plane")
    Call GetFreePlane(bISrcPlane, MYPLANE_LOCAL, idpDepthS16, False, "B (int) Source Plane")
    Call GetFreePlane(gCompPlane, MYPLANE_COMP_BAYER_G, idpDepthS16, False, "G Compressed Plane")
    Call GetFreePlane(bCompPlane, MYPLANE_COMP_BAYER_B, idpDepthS16, False, "B Compressed Plane")
    Call MultiMean(bayerPlaneR, "BAYER_R_ZONE2D", EEE_COLOR_FLAT, _
                   rISrcPlane, "LABZ2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, compFactor, compFactor)
    Call MultiMean(bayerPlaneG, "BAYER_G_ZONE2D", EEE_COLOR_FLAT, _
                   gCompPlane, "PCBAYER_G_Z2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, compFactor, compFactor)
    Call Copy(gCompPlane, "PCBAYER_G_Z2D", EEE_COLOR_FLAT, gISrcPlane, "LABZ2D", EEE_COLOR_FLAT)
    Call MultiMean(bayerPlaneB, "BAYER_B_ZONE2D", EEE_COLOR_FLAT, _
                   bCompPlane, "PCBAYER_B_Z2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, compFactor, compFactor)
    Call Copy(bCompPlane, "PCBAYER_B_Z2D", EEE_COLOR_FLAT, bISrcPlane, "LABZ2D", EEE_COLOR_FLAT)
    Call ReleasePlane(bayerPlaneR)
    Call ReleasePlane(bayerPlaneG)
    Call ReleasePlane(bayerPlaneB)
    Call ReleasePlane(gCompPlane)
    Call ReleasePlane(bCompPlane)
    
'直前のGとBのコピーは直接浮動小数にできないか見ること。
    'Flat Fielding
    Dim rFSrcPlane As CImgPlane
    Dim gFSrcPlane As CImgPlane
    Dim bFSrcPlane As CImgPlane
    Call GetFreePlane(rFSrcPlane, rISrcPlane.planeGroup, idpDepthF32, False, "R (Float) Source Plane")
    Call GetFreePlane(gFSrcPlane, gISrcPlane.planeGroup, idpDepthF32, False, "G (Float) Source Plane")
    Call GetFreePlane(bFSrcPlane, bISrcPlane.planeGroup, idpDepthF32, False, "B (Float) Source Plane")
    Call Copy(rISrcPlane, "LABZ2D", EEE_COLOR_FLAT, rFSrcPlane, "LABZ2D", EEE_COLOR_FLAT)
    Call Copy(gISrcPlane, "LABZ2D", EEE_COLOR_FLAT, gFSrcPlane, "LABZ2D", EEE_COLOR_FLAT)
    Call Copy(bISrcPlane, "LABZ2D", EEE_COLOR_FLAT, bFSrcPlane, "LABZ2D", EEE_COLOR_FLAT)
    Call ReleasePlane(rISrcPlane)
    Call ReleasePlane(gISrcPlane)
    Call ReleasePlane(bISrcPlane)
    
    With TheIDP.PlaneBank
        If .isExisting(PLANEBANK_FLATFIELD_R) Then Call .Delete(PLANEBANK_FLATFIELD_R)
        Call .Add(PLANEBANK_FLATFIELD_R, rFSrcPlane, True, True)
        
        If .isExisting(PLANEBANK_FLATFIELD_G) Then Call .Delete(PLANEBANK_FLATFIELD_G)
        Call .Add(PLANEBANK_FLATFIELD_G, gFSrcPlane, True, True)
        
        If .isExisting(PLANEBANK_FLATFIELD_B) Then Call .Delete(PLANEBANK_FLATFIELD_B)
        Call .Add(PLANEBANK_FLATFIELD_B, bFSrcPlane, True, True)
    End With
    
End Sub

Public Sub StdLab_GetMuraParameters()
    Dim MuraCol As Collection
    Set MuraCol = New Collection

    Call stdLab_GetMuraParam(MuraCol)
    
    Call StdLab_SetModParam(MuraCol)
    
    Set MuraCol = Nothing
End Sub

Private Sub stdLab_GetMuraParam(ByRef MuraCol As Collection)
    
    Dim Loopi As Long
    Dim LoopB As Long
    Dim buf As String
    Dim bufstr As String
    Dim StartLow As Long
    
    StartLow = 5
    
    Loopi = StartLow
    'C列　"Parameter Name" 検索
    Do Until Worksheets("Lab Mura Parameters").Range("C" & Loopi) = ""
        LoopB = Loopi
        
        'B列 "Mura Item" 検索
        Do While Worksheets("Lab Mura Parameters").Range("B" & LoopB) = ""
            LoopB = LoopB - 1
            
            If LoopB < StartLow Then
                Exit Do
            End If
        Loop
        'D列 "Value"取得 Keyは"Mura Item"&"Parameter Name"となる。
        MuraCol.Add Item:=Worksheets("Lab Mura Parameters").Range("D" & Loopi), key:=Worksheets("Lab Mura Parameters").Range("B" & LoopB) & Worksheets("Lab Mura Parameters").Range("C" & Loopi)

        Loopi = Loopi + 1
        If Loopi > 1000 Then
            Exit Do
        End If
    Loop

End Sub

Public Function StdLab_SetModParam(ByRef MuraCol As Collection)
  '#####  Set Lab Parameters  #####
    With MuraCol
        LDENL_FLAG_MEDIAN_TAPS = .Item("ConstantsLDENL_FLAG_MEDIAN_TAPS")
        LOCAL_DIFF_SIZE = .Item("ConstantsLOCAL_DIFF_SIZE")
        COMP_PRIMARY_TO_LOCAL = .Item("ConstantsCOMP_PRIMARY_TO_LOCAL")
        LABZ2D_X_SIZE = TheIDP.PMD("LABZ2D").width
        LABZ2D_Y_SIZE = TheIDP.PMD("LABZ2D").height
    End With

    LABZ2D_AREA = LABZ2D_X_SIZE * LABZ2D_Y_SIZE

    '#####  Set Kernel Parameters  #####
    With MuraCol
        LabKernel_LowPassH = .Item("Kernelskernel_LowPassH_ColorFloat")
        LabKernel_LowPassV = .Item("Kernelskernel_LowPassV_ColorFloat")
    End With

    '#####  Set Path Parameters  #####
    With MuraCol
        LabFlatFieldPath = .Item("PATHFlatFieldPath")
    End With

End Function

Public Function labProc_ReturnCompFactor() As Long
    labProc_ReturnCompFactor = COMP_PRIMARY_TO_LOCAL
End Function

Public Function labProc_ApplyLPF( _
    ByRef srcPlane As CImgPlane, _
    ByVal pZone As String, _
    ByVal kernelNameH As String, _
    ByVal kernelNameV As String)
    
    Const ZP_LEFT_IN As String = "LABZ2D_EX_LEFT_IN"
    Const ZP_RIGHT_IN As String = "LABZ2D_EX_RIGHT_IN"
    Const ZP_TOP_IN As String = "LABZ2D_EX_TOP_IN"
    Const ZP_BOTTOM_IN As String = "LABZ2D_EX_BOTTOM_IN"
    Const ZP_LEFT_OUT As String = "LABZ2D_EX_LEFT_OUT"
    Const ZP_RIGHT_OUT As String = "LABZ2D_EX_RIGHT_OUT"
    Const ZP_TOP_OUT As String = "LABZ2D_EX_TOP_OUT"
    Const ZP_BOTTOM_OUT As String = "LABZ2D_EX_BOTTOM_OUT"
    
    Dim workPlane As CImgPlane
    Call GetFreePlane(workPlane, srcPlane.planeGroup, idpDepthF32, False, "Work Plane")
    
    Call ExpandZoneEdgeValue7Bits(srcPlane, workPlane, ZP_LEFT_IN, ZP_LEFT_OUT)
    Call ExpandZoneEdgeValue7Bits(srcPlane, workPlane, ZP_RIGHT_IN, ZP_RIGHT_OUT)
    Call Convolution(srcPlane, pZone, EEE_COLOR_FLAT, workPlane, pZone, EEE_COLOR_FLAT, kernelNameH)
    Call ExpandZoneEdgeValue7Bits(workPlane, srcPlane, ZP_TOP_IN, ZP_TOP_OUT)
    Call ExpandZoneEdgeValue7Bits(workPlane, srcPlane, ZP_BOTTOM_IN, ZP_BOTTOM_OUT)
    Call Convolution(workPlane, pZone, EEE_COLOR_FLAT, srcPlane, pZone, EEE_COLOR_FLAT, kernelNameV)
    
    Call ReleasePlane(workPlane)
    
End Function

Public Sub ExpandZoneEdgeValue7Bits( _
    ByRef srcPlane As CImgPlane, _
    ByRef workPlane As CImgPlane, _
    ByVal inPlanePrefix As String, _
    ByVal outPlanePrefix As String)
    
    Call Copy(srcPlane, inPlanePrefix & "_00", EEE_COLOR_FLAT, _
              workPlane, outPlanePrefix & "_00", EEE_COLOR_FLAT)
    Call Copy(workPlane, inPlanePrefix & "_01", EEE_COLOR_FLAT, _
              srcPlane, outPlanePrefix & "_01", EEE_COLOR_FLAT)
                
    Call Copy(srcPlane, inPlanePrefix & "_02", EEE_COLOR_FLAT, _
              workPlane, outPlanePrefix & "_02", EEE_COLOR_FLAT)
    Call Copy(workPlane, inPlanePrefix & "_03", EEE_COLOR_FLAT, _
              srcPlane, outPlanePrefix & "_03", EEE_COLOR_FLAT)
    
    Call Copy(workPlane, inPlanePrefix & "_03", EEE_COLOR_FLAT, _
              srcPlane, outPlanePrefix & "_04", EEE_COLOR_FLAT)
End Sub

Public Sub procLab_GetFlatFields( _
    ByRef rPlane As CImgPlane, _
    ByRef gPlane As CImgPlane, _
    ByRef bPlane As CImgPlane)
    With TheIDP.PlaneBank
        Set rPlane = .Item(PLANEBANK_FLATFIELD_R)
        Set gPlane = .Item(PLANEBANK_FLATFIELD_G)
        Set bPlane = .Item(PLANEBANK_FLATFIELD_B)
    End With
End Sub

Public Sub labProc_RGB2LabDirect( _
    ByRef redPlane As CImgPlane, _
    ByRef greenPlane As CImgPlane, _
    ByRef bluePlane As CImgPlane, _
    ByRef lPlane As CImgPlane, _
    ByRef aPlane As CImgPlane, _
    ByRef bPlane As CImgPlane, _
    ByRef pZone As String, _
    ByRef wbZone As String, _
    ByRef resultLevel As Double)
    
    Const FACTOR_L_R As Double = 20.52852
    Const FACTOR_L_G As Double = 94.2384
    Const FACTOR_L_B As Double = 1.23308
    Const FACTOR_A_R As Double = 156.515
    Const FACTOR_A_G As Double = -251.2
    Const FACTOR_A_B As Double = 94.685
    Const FACTOR_B_R As Double = 35.394
    Const FACTOR_B_G As Double = 160.48
    Const FACTOR_B_B As Double = -195.874
    Dim site As Long
    
    '---- For white balance.
    Dim redMean(nSite) As Double
    Dim blueMean(nSite) As Double
    Dim greenMean(nSite) As Double
    Call Average(redPlane, wbZone, EEE_COLOR_FLAT, redMean)
    Call Average(greenPlane, wbZone, EEE_COLOR_FLAT, greenMean)
    Call Average(bluePlane, wbZone, EEE_COLOR_FLAT, blueMean)
    
    '---- RGB各画像に掛ける係数を求めます。
    Dim lrFactor(nSite) As Double
    Dim lgFactor(nSite) As Double
    Dim lbFactor(nSite) As Double
    Dim arFactor(nSite) As Double
    Dim agFactor(nSite) As Double
    Dim abFactor(nSite) As Double
    Dim brFactor(nSite) As Double
    Dim bgFactor(nSite) As Double
    Dim bbFactor(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            'L-16の画像生成用
            lrFactor(site) = Div(resultLevel, redMean(site), 0) * FACTOR_L_R
            lgFactor(site) = Div(resultLevel, greenMean(site), 0) * FACTOR_L_G
            lbFactor(site) = Div(resultLevel, blueMean(site), 0) * FACTOR_L_B
            
            'aの画像生成用
            arFactor(site) = Div(resultLevel, redMean(site), 0) * FACTOR_A_R
            agFactor(site) = Div(resultLevel, greenMean(site), 0) * FACTOR_A_G
            abFactor(site) = Div(resultLevel, blueMean(site), 0) * FACTOR_A_B
            
            'bの画像生成用
            brFactor(site) = Div(resultLevel, redMean(site), 0) * FACTOR_B_R
            bgFactor(site) = Div(resultLevel, greenMean(site), 0) * FACTOR_B_G
            bbFactor(site) = Div(resultLevel, blueMean(site), 0) * FACTOR_B_B
        End If
    Next site
    
    Dim workPlane1 As CImgPlane
    Dim workPlane2 As CImgPlane
    Call GetFreePlane(workPlane1, redPlane.planeGroup, idpDepthF32, False, "Work Plane 1")
    Call GetFreePlane(workPlane2, redPlane.planeGroup, idpDepthF32, False, "Work Plane 2")
    '---- L-16
    Call MultiplyConst(redPlane, pZone, EEE_COLOR_FLAT, lrFactor, lPlane, pZone, EEE_COLOR_FLAT)
    Call MultiplyConst(greenPlane, pZone, EEE_COLOR_FLAT, lgFactor, workPlane1, pZone, EEE_COLOR_FLAT)
    Call Add(lPlane, pZone, EEE_COLOR_FLAT, workPlane1, pZone, EEE_COLOR_FLAT, workPlane2, pZone, EEE_COLOR_FLAT)
    Call MultiplyConst(bluePlane, pZone, EEE_COLOR_FLAT, lbFactor, workPlane1, pZone, EEE_COLOR_FLAT)
    Call Add(workPlane2, pZone, EEE_COLOR_FLAT, workPlane1, pZone, EEE_COLOR_FLAT, lPlane, pZone, EEE_COLOR_FLAT)
    
    '---- A画像生成
    Call MultiplyConst(redPlane, pZone, EEE_COLOR_FLAT, arFactor, aPlane, pZone, EEE_COLOR_FLAT)
    Call MultiplyConst(greenPlane, pZone, EEE_COLOR_FLAT, agFactor, workPlane1, pZone, EEE_COLOR_FLAT)
    Call Add(aPlane, pZone, EEE_COLOR_FLAT, workPlane1, pZone, EEE_COLOR_FLAT, workPlane2, pZone, EEE_COLOR_FLAT)
    Call MultiplyConst(bluePlane, pZone, EEE_COLOR_FLAT, abFactor, workPlane1, pZone, EEE_COLOR_FLAT)
    Call Add(workPlane2, pZone, EEE_COLOR_FLAT, workPlane1, pZone, EEE_COLOR_FLAT, aPlane, pZone, EEE_COLOR_FLAT)
    
    '---- B画像生成
    Call MultiplyConst(redPlane, pZone, EEE_COLOR_FLAT, brFactor, bPlane, pZone, EEE_COLOR_FLAT)
    Call MultiplyConst(greenPlane, pZone, EEE_COLOR_FLAT, bgFactor, workPlane1, pZone, EEE_COLOR_FLAT)
    Call Add(bPlane, pZone, EEE_COLOR_FLAT, workPlane1, pZone, EEE_COLOR_FLAT, workPlane2, pZone, EEE_COLOR_FLAT)
    Call MultiplyConst(bluePlane, pZone, EEE_COLOR_FLAT, bbFactor, workPlane1, pZone, EEE_COLOR_FLAT)
    Call Add(workPlane2, pZone, EEE_COLOR_FLAT, workPlane1, pZone, EEE_COLOR_FLAT, bPlane, pZone, EEE_COLOR_FLAT)
    
    Call ReleasePlane(workPlane1)
    Call ReleasePlane(workPlane2)
    
End Sub

Public Function labProc_Ldenl( _
    ByRef vDifPlane As CImgPlane, _
    ByRef hDifPlane As CImgPlane, _
    ByVal sliceLevelLldif As Double, _
    ByVal nClusterEraserTaps As Long, _
    ByRef returnLdenl() As Double, _
    ByVal pixelCount As Long) As Boolean
    
    
    Dim returnResult(nSite) As Double
    Dim tmpCountRow(nSite) As Double
    Dim tmpCountCol(nSite) As Double
    Dim tmpColorResult(20) As Double
    
    Dim site As Long
    
    
    Dim FlagPlane As CImgPlane
    Dim flagMedPlane As CImgPlane
    Call GetFreePlane(FlagPlane, vDifPlane.planeGroup, idpDepthS16, True, "Flag Plane")    'lPlane??
    Call GetFreePlane(flagMedPlane, vDifPlane.planeGroup, idpDepthS16, True, "Flag Median Plane")
    
    '(+)
    Call vDifPlane.SetPMD("LABZ2D_LDENL_JUDGE_ROW")
    Call FlagPlane.SetPMD("LABZ2D_LDENL_JUDGE_ROW")
    Call TheHdw.IDP.Count(vDifPlane.Name, idpColorFlat, idpCountAbove, sliceLevelLldif, sliceLevelLldif, idpLimitInclude, , FlagPlane.Name, 1)
    Call Median(FlagPlane, "LABZ2D_LDENL_JUDGE_ROW", EEE_COLOR_FLAT, _
                flagMedPlane, "LABZ2D_LDENL_JUDGE_ROW", EEE_COLOR_FLAT, 1, nClusterEraserTaps)
    Call Median(flagMedPlane, "LABZ2D_LDENL_JUDGE_ROW", EEE_COLOR_FLAT, _
                FlagPlane, "LABZ2D_LDENL_JUDGE_ROW", EEE_COLOR_FLAT, nClusterEraserTaps, 1)
    With TheHdw.IDP
        .Accumulate vDifPlane.Name, idpColorFlat, idpAccumSum, , FlagPlane.Name, 1
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                .ReadResultFP site, vDifPlane.Name, EEE_COLOR_FLAT, idpReadSum, tmpColorResult
                If Abs(tmpColorResult(0)) < 1E+30 Then
                    returnLdenl(site) = tmpColorResult(0)
                Else
                    returnLdenl(site) = 0
                End If
            End If
        Next site
    End With
    
    '(-)
    Call vDifPlane.SetPMD("LABZ2D_LDENL_JUDGE_ROW")
    Call TheHdw.IDP.Count(vDifPlane.Name, idpColorFlat, idpCountBelow, _
                          -1 * sliceLevelLldif, -1 * sliceLevelLldif, idpLimitInclude, , FlagPlane.Name, 1)
    Call Median(FlagPlane, "LABZ2D_LDENL_JUDGE_ROW", EEE_COLOR_FLAT, _
                flagMedPlane, "LABZ2D_LDENL_JUDGE_ROW", EEE_COLOR_FLAT, 1, nClusterEraserTaps)
    Call Median(flagMedPlane, "LABZ2D_LDENL_JUDGE_ROW", EEE_COLOR_FLAT, _
                FlagPlane, "LABZ2D_LDENL_JUDGE_ROW", EEE_COLOR_FLAT, nClusterEraserTaps, 1)
    With TheHdw.IDP
        .Accumulate vDifPlane.Name, idpColorFlat, idpAccumSum, , FlagPlane.Name, 1
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                .ReadResultFP site, vDifPlane.Name, idpColorFlat, idpReadSum, tmpColorResult
                If (Abs(tmpColorResult(0)) < 1E+30) And (Abs(tmpColorResult(0)) > returnLdenl(site)) Then
                    returnLdenl(site) = Abs(tmpColorResult(0))
                End If
            End If
        Next site
    End With

    '-------------------------------------------------------------------------------
    ' H-Diff
    '-------------------------------------------------------------------------------
    Call WritePixel(FlagPlane, "LABZ2D_LOCAL_NOJUDGE_COL", EEE_COLOR_FLAT, 0)
    Call WritePixel(flagMedPlane, "LABZ2D_LOCAL_NOJUDGE_COL", EEE_COLOR_FLAT, 0)
    
    '(+)
    Call hDifPlane.SetPMD("LABZ2D_LDENL_JUDGE_COL")
    Call FlagPlane.SetPMD("LABZ2D_LDENL_JUDGE_COL")
    Call TheHdw.IDP.Count(hDifPlane.Name, idpColorFlat, idpCountAbove, sliceLevelLldif, sliceLevelLldif, idpLimitInclude, , FlagPlane.Name, 1)
    Call Median(FlagPlane, "LABZ2D_LDENL_JUDGE_COL", EEE_COLOR_FLAT, _
                flagMedPlane, "LABZ2D_LDENL_JUDGE_COL", EEE_COLOR_FLAT, 1, nClusterEraserTaps)
    Call Median(flagMedPlane, "LABZ2D_LDENL_JUDGE_COL", EEE_COLOR_FLAT, _
                FlagPlane, "LABZ2D_LDENL_JUDGE_COL", EEE_COLOR_FLAT, nClusterEraserTaps, 1)
    With TheHdw.IDP
        'Col(+)側体積導出
        .Accumulate hDifPlane.Name, idpColorFlat, idpAccumSum, , FlagPlane.Name, 1
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                .ReadResultFP site, hDifPlane.Name, idpColorFlat, idpReadSum, tmpColorResult
                If (Abs(tmpColorResult(0)) < 1E+30) And (tmpColorResult(0) > returnLdenl(site)) Then
                    returnLdenl(site) = tmpColorResult(0)
                End If
            End If
        Next site
    End With
    
    'フラグ作成(-)
    Call hDifPlane.SetPMD("LABZ2D_LDENL_JUDGE_COL")
    Call TheHdw.IDP.Count(hDifPlane.Name, idpColorFlat, idpCountBelow, -1 * sliceLevelLldif, -1 * sliceLevelLldif, idpLimitInclude, , FlagPlane.Name, 1)
    Call Median(FlagPlane, "LABZ2D_LDENL_JUDGE_COL", EEE_COLOR_FLAT, _
                flagMedPlane, "LABZ2D_LDENL_JUDGE_COL", EEE_COLOR_FLAT, 1, nClusterEraserTaps)
    Call Median(flagMedPlane, "LABZ2D_LDENL_JUDGE_COL", EEE_COLOR_FLAT, _
                FlagPlane, "LABZ2D_LDENL_JUDGE_COL", EEE_COLOR_FLAT, nClusterEraserTaps, 1)
    With TheHdw.IDP
        .Accumulate hDifPlane.Name, idpColorFlat, idpAccumSum, , FlagPlane.Name, 1
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                .ReadResultFP site, hDifPlane.Name, idpColorFlat, idpReadSum, tmpColorResult
                If (Abs(tmpColorResult(0)) < 1E+30) And (Abs(tmpColorResult(0)) > returnLdenl(site)) Then
                    returnLdenl(site) = Abs(tmpColorResult(0))
                End If
                returnLdenl(site) = returnLdenl(site) / pixelCount
            End If
        Next site
    End With
    
'    Call ReleasePlane(workPlane)
'    Call ReleasePlane(difPlane)
    Call ReleasePlane(FlagPlane)
    Call ReleasePlane(flagMedPlane)
    
End Function

Public Sub labProc_abmaxCol( _
    ByRef aPlane As CImgPlane, _
    ByRef bPlane As CImgPlane, _
    ByVal difZoneSource As String, _
    ByVal difZoneTarget As String, _
    ByVal judgeZone As String, _
    ByRef cmax() As Double)
    
    Dim site As Long
    Dim tmpColorResult(20) As Double
    
    Dim workPlane1 As CImgPlane
    Dim workPlane2 As CImgPlane
    Dim workPlane3 As CImgPlane
    Dim workplane4 As CImgPlane
    Call GetFreePlane(workPlane1, aPlane.planeGroup, idpDepthF32, False, "Work Plane 1")
    Call GetFreePlane(workPlane2, aPlane.planeGroup, idpDepthF32, False, "Work Plane 2")
    Call GetFreePlane(workPlane3, aPlane.planeGroup, idpDepthF32, False, "Work Plane 3")
    Call GetFreePlane(workplane4, aPlane.planeGroup, idpDepthF32, False, "Work Plane 4")

    '次にa画像のColumn差分→自乗を取ります。
    Call Copy(aPlane, difZoneTarget, EEE_COLOR_FLAT, workPlane3, difZoneTarget, EEE_COLOR_FLAT)
    Call Subtract(aPlane, difZoneSource, EEE_COLOR_FLAT, workPlane3, difZoneTarget, EEE_COLOR_FLAT, workPlane1, difZoneSource, EEE_COLOR_FLAT)
    Call Copy(workPlane1, judgeZone, EEE_COLOR_FLAT, workPlane2, judgeZone, EEE_COLOR_FLAT)
    Call Multiply(workPlane1, judgeZone, EEE_COLOR_FLAT, workPlane2, judgeZone, EEE_COLOR_FLAT, workplane4, judgeZone, EEE_COLOR_FLAT)
    'この時点で、workPlane4結果が入る
    
    '最後にb画像のColumn差分→自乗を取ります。
    Call Copy(bPlane, difZoneTarget, EEE_COLOR_FLAT, workPlane3, difZoneTarget, EEE_COLOR_FLAT)
    Call Subtract(bPlane, difZoneSource, EEE_COLOR_FLAT, workPlane3, difZoneTarget, EEE_COLOR_FLAT, workPlane1, difZoneSource, EEE_COLOR_FLAT)
    Call Copy(workPlane1, judgeZone, EEE_COLOR_FLAT, workPlane2, judgeZone, EEE_COLOR_FLAT)
    Call Multiply(workPlane1, judgeZone, EEE_COLOR_FLAT, workPlane2, judgeZone, EEE_COLOR_FLAT, workPlane3, judgeZone, EEE_COLOR_FLAT)
    Call Add(workPlane3, judgeZone, EEE_COLOR_FLAT, workplane4, judgeZone, EEE_COLOR_FLAT, workPlane1, judgeZone, EEE_COLOR_FLAT)
    'この時点で、workPlane1に結果が入る
    
    '最大値を取得して、Square Rootを取り、⊿E*の値を求めます。
    Dim returnMax(nSite) As Double
    Dim returnResult(nSite) As Double
    Call max(workPlane1, judgeZone, EEE_COLOR_FLAT, returnMax)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            cmax(site) = Sqr(returnMax(site))
        End If
    Next site
    
    Call ReleasePlane(workPlane1)
    Call ReleasePlane(workPlane2)
    Call ReleasePlane(workPlane3)
    Call ReleasePlane(workplane4)
    
End Sub

Public Sub labProc_abmaxRow( _
    ByRef aPlane As CImgPlane, _
    ByRef bPlane As CImgPlane, _
    ByVal difSize As Long, _
    ByVal pZone As String, _
    ByVal judgeZone As String, _
    ByRef cmax() As Double)
    
    Dim site As Long
    Dim tmpColorResult(20) As Double
    
    Dim workPlane1 As CImgPlane
    Dim workPlane2 As CImgPlane
    Dim workPlane3 As CImgPlane
    Dim workplane4 As CImgPlane
    Call GetFreePlane(workPlane1, aPlane.planeGroup, idpDepthF32, False, "Work Plane 1")
    Call GetFreePlane(workPlane2, aPlane.planeGroup, idpDepthF32, False, "Work Plane 2")
    Call GetFreePlane(workPlane3, aPlane.planeGroup, idpDepthF32, False, "Work Plane 3")
    Call GetFreePlane(workplane4, aPlane.planeGroup, idpDepthF32, False, "Work Plane 4")

    '次にa画像のRow差分→自乗を取ります。
    Call SubRows(aPlane, pZone, EEE_COLOR_FLAT, workPlane1, pZone, EEE_COLOR_FLAT, difSize)
    Call Copy(workPlane1, judgeZone, EEE_COLOR_FLAT, workPlane2, judgeZone, EEE_COLOR_FLAT)
    Call Multiply(workPlane1, judgeZone, EEE_COLOR_FLAT, workPlane2, judgeZone, EEE_COLOR_FLAT, workPlane3, judgeZone, EEE_COLOR_FLAT)
    'この時点で、workPlane3に結果が入る
    
    Call SubRows(bPlane, pZone, EEE_COLOR_FLAT, workPlane1, pZone, EEE_COLOR_FLAT, difSize)
    Call Copy(workPlane1, judgeZone, EEE_COLOR_FLAT, workPlane2, judgeZone, EEE_COLOR_FLAT)
    Call Multiply(workPlane1, judgeZone, EEE_COLOR_FLAT, workPlane2, judgeZone, EEE_COLOR_FLAT, workplane4, judgeZone, EEE_COLOR_FLAT)
    Call Add(workPlane3, judgeZone, EEE_COLOR_FLAT, workplane4, judgeZone, EEE_COLOR_FLAT, workPlane1, judgeZone, EEE_COLOR_FLAT)
    'この時点で、workPlane1に結果が入る。
    
    '最大値を取得して、Square Rootを取り、sqr(a^2 + b^2)の値を求めます。
    Dim returnMax(nSite) As Double
    Call max(workPlane1, judgeZone, EEE_COLOR_FLAT, returnMax)
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            cmax(site) = Sqr(returnMax(site))
        End If
    Next site
    
    Call ReleasePlane(workPlane1)
    Call ReleasePlane(workPlane2)
    Call ReleasePlane(workPlane3)
    Call ReleasePlane(workplane4)
    
End Sub

Public Function Lab_Haki(ByRef lPlane As CImgPlane, ByVal sliceLevelLldif As Double, ByRef returnLdenl() As Double)

    Const PROCESS_ZONE As String = "LABZ2D"
    Const HDIF_ZONE_SOURCE As String = "LABZ2D_LOCAL_SOURCE_COL"
    Const HDIF_ZONE_TARGET As String = "LABZ2D_LOCAL_TARGET_COL"
    
    Const HDIF_LDENL_JUDGE As String = "LABZ2D_LDENL_JUDGE_COL"
    Const VDIF_LDENL_JUDGE As String = "LABZ2D_LDENL_JUDGE_ROW"
    
    Const HDIF_VDIF_ZONE_NOJUDGE As String = "LABZ2D_LOCAL_NOJUDGE_COL"
    
    Dim tmpColorResult(20) As Double
        
    Dim site As Long
        
    Dim difPlane As CImgPlane 'f4vmcu06
    Call GetFreePlane(difPlane, lPlane.planeGroup, idpDepthF32, False, "Diff Plane")
    
    Dim FlagPlane As CImgPlane
    Dim flagMedPlane As CImgPlane
    Call GetFreePlane(FlagPlane, lPlane.planeGroup, idpDepthS16, True, "Flag Plane")
    Call GetFreePlane(flagMedPlane, lPlane.planeGroup, idpDepthS16, True, "Flag Median Plane")

    '-------------------------------------------------------------------------------
    ' V-Diff
    '-------------------------------------------------------------------------------
    Call SubRows(lPlane, PROCESS_ZONE, EEE_COLOR_FLAT, difPlane, PROCESS_ZONE, EEE_COLOR_FLAT, LOCAL_DIFF_SIZE)
    
    '(+)
    Call difPlane.SetPMD(VDIF_LDENL_JUDGE)
    Call FlagPlane.SetPMD(VDIF_LDENL_JUDGE)
    Call TheHdw.IDP.Count(difPlane.Name, idpColorFlat, idpCountAbove, sliceLevelLldif, sliceLevelLldif, idpLimitInclude, , FlagPlane.Name, 1)
    Call Median(FlagPlane, VDIF_LDENL_JUDGE, EEE_COLOR_FLAT, _
                flagMedPlane, VDIF_LDENL_JUDGE, EEE_COLOR_FLAT, 1, LDENL_FLAG_MEDIAN_TAPS)
    Call Median(flagMedPlane, VDIF_LDENL_JUDGE, EEE_COLOR_FLAT, _
                FlagPlane, VDIF_LDENL_JUDGE, EEE_COLOR_FLAT, LDENL_FLAG_MEDIAN_TAPS, 1)
    With TheHdw.IDP
        .Accumulate difPlane.Name, idpColorFlat, idpAccumSum, , FlagPlane.Name, 1
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                .ReadResultFP site, difPlane.Name, idpColorFlat, idpReadSum, tmpColorResult
                If Abs(tmpColorResult(0)) < 1E+30 Then
                    returnLdenl(site) = tmpColorResult(0)
                Else
                    returnLdenl(site) = 0
                End If
            End If
        Next site
    End With
    
    '(-)
    Call difPlane.SetPMD(VDIF_LDENL_JUDGE)
    Call TheHdw.IDP.Count(difPlane.Name, idpColorFlat, idpCountBelow, _
                          -1 * sliceLevelLldif, -1 * sliceLevelLldif, idpLimitInclude, , FlagPlane.Name, 1)
    Call Median(FlagPlane, VDIF_LDENL_JUDGE, EEE_COLOR_FLAT, _
                flagMedPlane, VDIF_LDENL_JUDGE, EEE_COLOR_FLAT, 1, LDENL_FLAG_MEDIAN_TAPS)
    Call Median(flagMedPlane, VDIF_LDENL_JUDGE, EEE_COLOR_FLAT, _
                FlagPlane, VDIF_LDENL_JUDGE, EEE_COLOR_FLAT, LDENL_FLAG_MEDIAN_TAPS, 1)
    With TheHdw.IDP
        .Accumulate difPlane.Name, idpColorFlat, idpAccumSum, , FlagPlane.Name, 1
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                .ReadResultFP site, difPlane.Name, idpColorFlat, idpReadSum, tmpColorResult
                If (Abs(tmpColorResult(0)) < 1E+30) And (Abs(tmpColorResult(0)) > returnLdenl(site)) Then
                    returnLdenl(site) = Abs(tmpColorResult(0))
                End If
            End If
        Next site
    End With

    '-------------------------------------------------------------------------------
    ' H-Diff
    '-------------------------------------------------------------------------------
    Call WritePixel(FlagPlane, HDIF_VDIF_ZONE_NOJUDGE, EEE_COLOR_FLAT, 0)
    Call WritePixel(flagMedPlane, HDIF_VDIF_ZONE_NOJUDGE, EEE_COLOR_FLAT, 0)
    
    Dim workPlane As CImgPlane
    Call GetFreePlane(workPlane, lPlane.planeGroup, idpDepthF32, False, "Work Plane")
    Call Copy(lPlane, PROCESS_ZONE, EEE_COLOR_FLAT, workPlane, PROCESS_ZONE, EEE_COLOR_FLAT)
    Call Subtract(lPlane, HDIF_ZONE_SOURCE, EEE_COLOR_FLAT, _
                  workPlane, HDIF_ZONE_TARGET, EEE_COLOR_FLAT, _
                  difPlane, HDIF_ZONE_SOURCE, EEE_COLOR_FLAT)
                  
    '(+)
    Call difPlane.SetPMD(HDIF_LDENL_JUDGE)
    Call FlagPlane.SetPMD(HDIF_LDENL_JUDGE)
    Call TheHdw.IDP.Count(difPlane.Name, idpColorFlat, idpCountAbove, sliceLevelLldif, sliceLevelLldif, idpLimitInclude, , FlagPlane.Name, 1)
    Call Median(FlagPlane, HDIF_LDENL_JUDGE, EEE_COLOR_FLAT, _
                flagMedPlane, HDIF_LDENL_JUDGE, EEE_COLOR_FLAT, 1, LDENL_FLAG_MEDIAN_TAPS)
    Call Median(flagMedPlane, HDIF_LDENL_JUDGE, EEE_COLOR_FLAT, _
                FlagPlane, HDIF_LDENL_JUDGE, EEE_COLOR_FLAT, LDENL_FLAG_MEDIAN_TAPS, 1)
    With TheHdw.IDP
        'Col(+)側体積導出
        .Accumulate difPlane.Name, idpColorFlat, idpAccumSum, , FlagPlane.Name, 1
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                .ReadResultFP site, difPlane.Name, idpColorFlat, idpReadSum, tmpColorResult
                If (Abs(tmpColorResult(0)) < 1E+30) And (tmpColorResult(0) > returnLdenl(site)) Then
                    returnLdenl(site) = tmpColorResult(0)
                End If
            End If
        Next site
    End With
    
    'フラグ作成(-)
    Call difPlane.SetPMD(HDIF_LDENL_JUDGE)
    Call TheHdw.IDP.Count(difPlane.Name, idpColorFlat, idpCountBelow, -1 * sliceLevelLldif, -1 * sliceLevelLldif, idpLimitInclude, , FlagPlane.Name, 1)
    Call Median(FlagPlane, HDIF_LDENL_JUDGE, EEE_COLOR_FLAT, _
                flagMedPlane, HDIF_LDENL_JUDGE, EEE_COLOR_FLAT, 1, LDENL_FLAG_MEDIAN_TAPS)
    Call Median(flagMedPlane, HDIF_LDENL_JUDGE, EEE_COLOR_FLAT, _
                FlagPlane, HDIF_LDENL_JUDGE, EEE_COLOR_FLAT, LDENL_FLAG_MEDIAN_TAPS, 1)
    With TheHdw.IDP
        .Accumulate difPlane.Name, idpColorFlat, idpAccumSum, , FlagPlane.Name, 1
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                .ReadResultFP site, difPlane.Name, idpColorFlat, idpReadSum, tmpColorResult
                If (Abs(tmpColorResult(0)) < 1E+30) And (Abs(tmpColorResult(0)) > returnLdenl(site)) Then
                    returnLdenl(site) = Abs(tmpColorResult(0))
                End If
                returnLdenl(site) = returnLdenl(site) / LABZ2D_AREA
            End If
        Next site
    End With
    
    Call ReleasePlane(workPlane)
    Call ReleasePlane(difPlane)
    Call ReleasePlane(FlagPlane)
    Call ReleasePlane(flagMedPlane)

End Function
'========================================================================================================================================

Public Sub StdCM_Initialize(ByVal pType As String, ByVal ZONE_FULL As Variant, ByVal ZONE_ZONE3 As Variant, ByVal clampZone As Variant, ByVal paramSheetName As String)

    'Local variables
    Dim site As Long                'For site loop

    Dim Flg_Active(nSite) As Long

    If TheIDP.KernelManager.IsExist("kernel_StdCM3x3") = True Then Exit Sub  'CHECK

     'ALL SITE ACTIVE
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Flg_Active(site) = 1
        End If
    Next site
    TheExec.sites.SetAllActive (True)
    
    If Not TheIDP.LUTManager.IsExist("lut_arctan") Then Call lut_v30

    'To read parameters
    Call StdCM_GetMuraParameters(paramSheetName)
     'To define kernels.
    Call ker_Color_v30
    'Flat Field Mode
    If IsClsMuraFlatFieldingOn Then
        
        Call CM_MakeFlatFieldImagePlane(pType, ZONE_FULL, ZONE_ZONE3, clampZone)
        
    Else
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                refFlatFieldMeanR(site) = 1
                RefFlatFieldMeanG(site) = 1
                refFlatFieldMeanB(site) = 1
            End If
        Next site
    End If

    'SITE ACTIVE RETURN
    For site = 0 To nSite
        If Flg_Active(site) = 0 Then
            TheExec.sites.site(site).Active = False
        End If
    Next site
    
End Sub

Private Sub CM_MakeFlatFieldImagePlane(ByVal pType As String, ByVal ZONE_FULL As Variant, ByVal ZONE_ZONE3 As Variant, ByVal clampZone As Variant)

    '[Edit Here] Filename Prefix Setting
    Const FF_FILE_PREFIX As String = "HL-"
    
    Dim site As Long                'For site loop
    Dim Bclamp(nSite) As Double     'For opb clamp
    
    'To read flat field image.
    Dim rawPlane As CImgPlane
    Call GetFreePlane(rawPlane, pType, idpDepthS16, False, "Flat Field Input")
  
    Dim myTesterName As String
    myTesterName = Format(Sw_Node, "000")
  
    Dim myTesterSite As String
    Dim myFileName As String
  
    For site = 0 To nSite
        myTesterSite = "-Site-" & Format(site, "0")
        myFileName = FF_FILE_PREFIX & myTesterName & myTesterSite & ".stb"
        Call InPutImage(site, rawPlane, ZONE_FULL, FlatFieldPath & myTesterName & "\" & myFileName)
    Next site
        
    'OPB Clamp for Flat Field image.
    Call Average(rawPlane, clampZone, EEE_COLOR_FLAT, Bclamp)
    Dim workPlane1 As CImgPlane
    Call GetFreePlane(workPlane1, rawPlane.planeGroup, idpDepthS16, , "workPlane1")
    Call SubtractConst(rawPlane, ZONE_ZONE3, EEE_COLOR_FLAT, Bclamp, workPlane1, ZONE_ZONE3, EEE_COLOR_FLAT)
    
    'To perform noise reduction with median filter for flat field image.
    Dim workPlane2 As CImgPlane
    Call GetFreePlane(workPlane2, rawPlane.planeGroup, idpDepthS16, False, "workPlane2")
    Call Median(workPlane1, ZONE_ZONE3, EEE_COLOR_ALL, workPlane2, ZONE_ZONE3, EEE_COLOR_ALL, 5, 1)
    Call Median(workPlane2, ZONE_ZONE3, EEE_COLOR_ALL, workPlane1, ZONE_ZONE3, EEE_COLOR_ALL, 1, 5)
    Call ReleasePlane(workPlane2)
    
    'Bayer separation for flat field image.
    Dim bayerRedPlane As CImgPlane
    Dim bayerGreenPlane As CImgPlane
    Dim bayerBluePlane As CImgPlane
    Call GetFreePlane(bayerRedPlane, "rbayer", idpDepthS16, False, "redPlane for bayer")
    Call GetFreePlane(bayerGreenPlane, "gbayer", idpDepthS16, False, "greenPlane for bayer")
    Call GetFreePlane(bayerBluePlane, "bbayer", idpDepthS16, False, "bluePlane for bayer")

    Dim tmpBayerPlane As CImgPlane
    Call GetFreePlane(tmpBayerPlane, "allbayer", idpDepthS16, False, "Clamp Image (Bayer plane)")
    Call Copy(workPlane1, "ZONE2D_BAYER", EEE_COLOR_FLAT, tmpBayerPlane, "ALLBAYER_ZONE2D", EEE_COLOR_FLAT)
    If StdCM_ColorMapName = "Bayer" Then
            Call StdCM_SeparateRGBforBayer(tmpBayerPlane, "ALLBAYER_ZONE2D", _
                                           bayerRedPlane, "RBAYER_FULL", _
                                           bayerGreenPlane, "GBAYER_FULL", _
                                           bayerBluePlane, "BBAYER_FULL")
    ElseIf StdCM_ColorMapName = "ClearVid-Yoko" Then
            'In case of Clear-Vid (Yoko-Tsubushi), use the following, instead.
            Call StdCM_SeparateRGBforClearVid(tmpBayerPlane, "ALLBAYER_ZONE2D", _
                                              bayerRedPlane, "RBAYER_FULL", _
                                              bayerGreenPlane, "GBAYER_FULL", _
                                              bayerBluePlane, "BBAYER_FULL")
    ElseIf StdCM_ColorMapName = "RGBandIR" Then
            'In case of Bayer with "Gb" replaced with IR or something that is to be ignored.
            Call StdCM_SeparateRGBandIRforBayer(tmpBayerPlane, "ALLBAYER_ZONE2D", _
                                                bayerRedPlane, "RBAYER_FULL", _
                                                bayerGreenPlane, "GBAYER_FULL", _
                                                bayerBluePlane, "BBAYER_FULL")
    Else
            'No Bayer, No Clear-Vid (YokoTsubushi), please create bayer separation function here.
    End If
    Call ReleasePlane(tmpBayerPlane)
    Call ReleasePlane(workPlane1)

    'To copy R/G/B flat field images to "YLINE" planes
    Dim rRawPlane As CImgPlane
    Dim gRawPlane As CImgPlane
    Dim bRawPlane As CImgPlane
    Call GetFreePlane(rRawPlane, "pyline", idpDepthF32, False, "Red Flat Field")
    Call GetFreePlane(gRawPlane, "pyline", idpDepthF32, False, "Green Flat Field")
    Call GetFreePlane(bRawPlane, "pyline", idpDepthF32, False, "Blue Flat Field")
    Call Copy(bayerRedPlane, "RBAYER_FULL", EEE_COLOR_FLAT, rRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    Call Copy(bayerGreenPlane, "GBAYER_FULL", EEE_COLOR_FLAT, gRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    Call Copy(bayerBluePlane, "BBAYER_FULL", EEE_COLOR_FLAT, bRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    Call ReleasePlane(bayerRedPlane)
    Call ReleasePlane(bayerGreenPlane)
    Call ReleasePlane(bayerBluePlane)

    'To calculate mean value of each flat field.
    Call Average(rRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, refFlatFieldMeanR)
    Call Average(gRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, RefFlatFieldMeanG)
    Call Average(bRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, refFlatFieldMeanB)

    'To deposit R/G/B flat field images to the Plane Bank.
    Call TheIDP.PlaneBank.Add("FLAT FIELD RED", rRawPlane, True, True)
    Call TheIDP.PlaneBank.Add("FLAT FIELD GREEN", gRawPlane, True, True)
    Call TheIDP.PlaneBank.Add("FLAT FIELD BLUE", bRawPlane, True, True)

    Call ReleasePlane(rawPlane)

End Sub

Public Sub stdCM_GetFFMean( _
    ByRef meanR() As Double, _
    ByRef meanG() As Double, _
    ByRef meanB() As Double)
    meanR = refFlatFieldMeanR
    meanG = RefFlatFieldMeanG
    meanB = refFlatFieldMeanB
End Sub

Private Sub ker_Color_v30()

    With TheIDP
        .CreateKernel LowPassFilterNameH, idpKernelFloat, 15, 1, 0, LpfKernel
        .CreateKernel LowPassFilterNameV, idpKernelFloat, 1, 15, 0, LpfKernel
        
        .CreateKernel "kernel_StdCM3x3", idpKernelInteger, 3, 3, 0, "1 1 1 1 1 1 1 1 1"
'                               1   1   1
'                               1   1   1
'                               1   1   1

        .CreateKernel "kernel_StdCM5x5", idpKernelInteger, 5, 5, 0, "1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1"
'                               1   1   1   1   1
'                               1   1   1   1   1
'                               1   1   1   1   1
'                               1   1   1   1   1
'                               1   1   1   1   1
        .CreateKernel "kernel_Ygloblal2", idpKernelInteger, 9, 7, 0, "1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1"
'                             TAP:X9 , Y7
    End With
    
End Sub

Public Sub lut_v30()

'{
'   PMD設定、Kernel設定と一緒に管理した方がいい。
'   TheIDP.RemoveResourcesでそれらとまとめてクリアされるので。
'   名前はやりたい内容を表すものに変更。
'}
    Dim intLoopCount As Long
    Dim lngOutVal As Long

'   /****** [arctan] *****/
    TheIDP.CreateIndexLUT "lut_arctan", -32767, -16384, -90, -90, 16
    For intLoopCount = -16383 To 16383 Step 1
        lngOutVal = Int(Atn(CDbl(intLoopCount) / 128#) * 180# / pi + 0.5)
        TheIDP.CreateIndexLUT "lut_arctan", intLoopCount, intLoopCount + 1, lngOutVal, lngOutVal, 16
    Next intLoopCount
    TheIDP.CreateIndexLUT "lut_arctan", 16384, 32766, 90, 90, 16

'   /****** [15] *****/
    With TheIDP                                                             ' Look Up Table 15
        .CreateIndexLUT "lut_15_new", -2048, -91, 0, 0, 12
        .CreateIndexLUT "lut_15_new", -90, -90, 256, 256, 12
        .CreateIndexLUT "lut_15_new", -89, -85, 250, 250, 12
        .CreateIndexLUT "lut_15_new", -84, -75, 245, 245, 12
        .CreateIndexLUT "lut_15_new", -74, -60, 240, 240, 12
        .CreateIndexLUT "lut_15_new", -59, -50, 235, 235, 12
        .CreateIndexLUT "lut_15_new", -49, -40, 230, 230, 12
        .CreateIndexLUT "lut_15_new", -39, -35, 226, 226, 12
        .CreateIndexLUT "lut_15_new", -34, -30, 222, 222, 12
        .CreateIndexLUT "lut_15_new", -29, -25, 217, 217, 12
        .CreateIndexLUT "lut_15_new", -24, -20, 213, 213, 12
        .CreateIndexLUT "lut_15_new", -19, -15, 202, 202, 12
        .CreateIndexLUT "lut_15_new", -14, -10, 199, 199, 12
        .CreateIndexLUT "lut_15_new", -9, -5, 195, 195, 12
        .CreateIndexLUT "lut_15_new", -4, 0, 192, 192, 12
        .CreateIndexLUT "lut_15_new", 0, 5, 192, 192, 12
        .CreateIndexLUT "lut_15_new", 6, 10, 195, 195, 12
        .CreateIndexLUT "lut_15_new", 11, 15, 202, 202, 12
        .CreateIndexLUT "lut_15_new", 16, 20, 206, 206, 12
        .CreateIndexLUT "lut_15_new", 21, 25, 217, 217, 12
        .CreateIndexLUT "lut_15_new", 26, 30, 222, 222, 12
        .CreateIndexLUT "lut_15_new", 31, 35, 230, 230, 12
        .CreateIndexLUT "lut_15_new", 36, 40, 235, 235, 12
        .CreateIndexLUT "lut_15_new", 41, 45, 240, 240, 12
        .CreateIndexLUT "lut_15_new", 46, 55, 245, 245, 12
        .CreateIndexLUT "lut_15_new", 56, 80, 250, 250, 12
        .CreateIndexLUT "lut_15_new", 81, 90, 245, 245, 12
        .CreateIndexLUT "lut_15_new", 91, 100, 240, 240, 12
        .CreateIndexLUT "lut_15_new", 101, 115, 235, 235, 12
        .CreateIndexLUT "lut_15_new", 116, 125, 240, 240, 12
        .CreateIndexLUT "lut_15_new", 126, 135, 245, 245, 12
        .CreateIndexLUT "lut_15_new", 136, 145, 250, 250, 12
        .CreateIndexLUT "lut_15_new", 146, 150, 256, 256, 12
        .CreateIndexLUT "lut_15_new", 151, 165, 262, 262, 12
        .CreateIndexLUT "lut_15_new", 166, 175, 268, 268, 12
        .CreateIndexLUT "lut_15_new", 176, 185, 274, 274, 12
        .CreateIndexLUT "lut_15_new", 186, 195, 281, 281, 12
        .CreateIndexLUT "lut_15_new", 196, 235, 288, 288, 12
        .CreateIndexLUT "lut_15_new", 236, 240, 281, 281, 12
        .CreateIndexLUT "lut_15_new", 241, 245, 274, 274, 12
        .CreateIndexLUT "lut_15_new", 246, 255, 268, 268, 12
        .CreateIndexLUT "lut_15_new", 256, 260, 262, 262, 12
        .CreateIndexLUT "lut_15_new", 261, 270, 256, 256, 12
        .CreateIndexLUT "lut_15_new", 271, 275, 250, 250, 12
        .CreateIndexLUT "lut_15_new", 276, 285, 245, 245, 12
        .CreateIndexLUT "lut_15_new", 286, 300, 240, 240, 12
        .CreateIndexLUT "lut_15_new", 301, 310, 235, 235, 12
        .CreateIndexLUT "lut_15_new", 311, 320, 230, 230, 12
        .CreateIndexLUT "lut_15_new", 321, 325, 226, 226, 12
        .CreateIndexLUT "lut_15_new", 326, 330, 222, 222, 12
        .CreateIndexLUT "lut_15_new", 331, 335, 217, 217, 12
        .CreateIndexLUT "lut_15_new", 336, 340, 213, 213, 12
        .CreateIndexLUT "lut_15_new", 341, 345, 202, 202, 12
        .CreateIndexLUT "lut_15_new", 346, 350, 199, 199, 12
        .CreateIndexLUT "lut_15_new", 351, 355, 195, 195, 12
        .CreateIndexLUT "lut_15_new", 356, 360, 192, 192, 12
        .CreateIndexLUT "lut_15_new", 361, 2047, 0, 0, 12
    End With


End Sub

Public Sub StdCM_GetMuraParameters(ByVal paramSheetName As String)
    Dim MuraCol As Collection
    Set MuraCol = New Collection

    Call stdCM_GetMuraParam(MuraCol, paramSheetName)
    
    Call stdCM_SetCommonParam(MuraCol)
    Call StdCM_SetModParam(MuraCol)
    
    Set MuraCol = Nothing
End Sub

Private Sub stdCM_GetMuraParam(ByRef MuraCol As Collection, ByVal paramSheetName As String)
    
    Dim Loopi As Long
    Dim LoopB As Long
    Dim buf As String
    Dim bufstr As String
    Dim StartLow As Long
    
    StartLow = 5
    
    Loopi = StartLow
    'C列　"Parameter Name" 検索
    Do Until Worksheets(paramSheetName).Range("C" & Loopi) = ""
        LoopB = Loopi
        
        'B列 "Mura Item" 検索
        Do While Worksheets(paramSheetName).Range("B" & LoopB) = ""
            LoopB = LoopB - 1
            
            If LoopB < StartLow Then
                Exit Do
            End If
        Loop
        'D列 "Value"取得 Keyは"Mura Item"&"Parameter Name"となる。
On Error GoTo ADDITION_FAILED
        MuraCol.Add Item:=Worksheets(paramSheetName).Range("D" & Loopi), key:=Worksheets(paramSheetName).Range("B" & LoopB) & Worksheets(paramSheetName).Range("C" & Loopi)
On Error GoTo 0

        Loopi = Loopi + 1
        If Loopi > 1000 Then
            Exit Do
        End If
    Loop
    Exit Sub
    
ADDITION_FAILED:
    
End Sub

Private Function stdCM_SetCommonParam(ByRef MuraCol As Collection)

  '#####  Set Common Parameters  #####
    With MuraCol
        MagicNumber = .Item("ConstantsMagicNumber")
        RyGain = .Item("ConstantsRyGain")
        ByGain = .Item("ConstantsByGain")
        RyHue = .Item("ConstantsRyHue")
        ByHue = .Item("ConstantsByHue")
        IsClsMuraFlatFieldingOn = .Item("ConstantsIsFlatFieldingOn")
    End With

    '#####  Set Kernel Parameters  #####

    With MuraCol
        LowPassFilterNameH = .Item("KernelsLowPassH")
        LowPassFilterNameV = .Item("KernelsLowPassV")
    End With

    '#####  Set LUT Parameters  #####

    With MuraCol
        LookUpTableArctan = .Item("LUTArcTangent")
    End With
    
    '#####  Set Path Parameters  #####
    With MuraCol
        FlatFieldPath = .Item("PATHFlatFieldPath")
    End With

    With MuraCol
        LpfKernel = .Item("KernelsLowPassFilterKernel")
    End With
End Function

Public Function StdCM_SetModParam(ByRef MuraCol As Collection)
 '#####  Set Common Parameters  #####
    With MuraCol
        UnitSize = .Item("CommonUnitSize")
    End With
    
    '#####  Set Common Parameters  #####
    With MuraCol
        MINIMUM_DIF_SIZE_REQUIRED = .Item("ConstantsMinimumDefferentialSizeAllowed")
        SIZE_RATIO_GLOBAL_OVER_LOCAL = .Item("ConstantsSizeRatioGlobalOverLocal")

        COMP_RGB_TO_Y = .Item("ConstantsCompressionSize_RGBtoYLINE")
        COMP_YLINE_TO_YLOCAL = .Item("ConstantsCompressionSize_YLINEtoYLOCAL")
        COMP_YLOCAL_TO_YGLOBAL = .Item("ConstantsCompressionSize_YLOCALtoYGLOBAL")
        COMP_YLOCAL_TO_CLOCAL = .Item("ConstantsCompressionSize_YLOCALtoCLOCAL")
        COMP_CLOCAL_TO_CGLOBAL = .Item("ConstantsCompressionSize_CLOCALtoCGLOBAL")
        
        colorMapName = .Item("ConstantsBaseColorMapName")
    End With

    '#####  Set Kernel Parameters  #####
    With MuraCol
        Kernel_LowPassH = .Item("KernelsLowPassH")
        Kernel_LowPassV = .Item("KernelsLowPassV")
        Kernel_YLine = .Item("KernelsYLine")
        Kernel_Ylocal = .Item("KernelsYLocal")
        Kernel_YFrame2D = .Item("KernelsYFrame2D")
        Kernel_YFrame2 = .Item("KernelsYFrame2")
        Kernel_YGlobal = .Item("KernelsYGlobal")
'        Kernel_YGlobal2 = .Item("KernelsYGlobal2")
   End With

'    '#####  Set Slice Parameters  #####
'    With MuraCol
'        yLineSlice = .Item("SliceLevelsYLine")
'        yLocalSlice = .Item("SliceLevelsYLocal")
'        yFrame2DSlice = .Item("SliceLevelsYFrame2D")
'        yFrame2Slice = .Item("SliceLevelsYFrame2")
'        yGlobalSlice = .Item("SliceLevelsYGlobal")
'    End With

    '#####  Set YLINE Parameters  #####
    With MuraCol
        yLineDif = .Item("YLINEDifferentialDistance")
        yLineCoef = .Item("YLINECoefficient")
        yLineCompPix = .Item("YLINECompressPixel")
        yLineCount = .Item("YLINEFlagSlice")
    End With
        
    '#####  Set YLOCAL Parameters  #####
    With MuraCol
        yLocalDif = .Item("YLOCALDifferentialDistance")
        yLocalCoef = .Item("YLOCALCoefficient")
        yLocalCount = .Item("YLOCALFlagSlice")
    End With
    
    '#####  Set YGLOBAL Parameters  #####
    With MuraCol
        yGlobDif = .Item("YGLOBALDifferentialDistance")
        yGlobCoef = .Item("YGLOBALCoefficient")
        yGlobCount = .Item("YGLOBALFlagSlice")
    End With
    
    '#####  Set YGLOBAL2 Parameters  #####
    With MuraCol
'        yGlob2HDif = .Item("YGLOBAL2DifferentialDistanceH")
'        yGlob2VDif = .Item("YGLOBAL2DifferentialDistanceV")
'        yGlob2Count = .Item("YGLOBAL2FlagSlice")
    End With

    '#####  Set CLOCAL Parameters  #####
    With MuraCol
        cLocalDif = .Item("CLOCALDifferentialDistance")
        cLocalCoef = .Item("CLOCALCoefficient")
    End With
      
    '#####  Set CGLOBAL Parameters  #####
    With MuraCol
        cGlobDif = .Item("CGLOBALDifferentialDistance")
        cGlobCoef = .Item("CGLOBALCoefficient")
    End With
        
    '#####  Set LUT Parameters  #####
    With MuraCol
        lookUpTable_lut15 = .Item("LUTHueGain")
    End With

End Function

Private Function StdCM_CompRGB2Y() As Long
    StdCM_CompRGB2Y = COMP_RGB_TO_Y
End Function

Private Function StdCM_ColorMapName() As String
    StdCM_ColorMapName = colorMapName
End Function

Public Sub StdCM_RegisteredValue( _
        ByVal pRegisterdName As String, ByRef pRegisterdValue() As Double)

    Call TheResult.Add(pRegisterdName, pRegisterdValue)
End Sub

Public Function StdCM_GetMax(ParamArray dataArr() As Variant) As Double

    If UBound(dataArr) < 0 Then
        StdCM_GetMax = 0
        Exit Function
    End If

    Dim i As Long
    Dim tmp As Double
    tmp = dataArr(0)
    For i = 1 To UBound(dataArr)
        If tmp < dataArr(i) Then
            tmp = dataArr(i)
        End If
    Next i
    StdCM_GetMax = tmp

End Function

Public Sub StdCM_SeparateRGBforBayer( _
    ByRef srcPlane As CImgPlane, _
    ByVal srcZone As String, _
    ByRef redPlane As CImgPlane, _
    ByVal redZone As String, _
    ByRef greenPlane As CImgPlane, _
    ByVal greenZone As String, _
    ByRef bluePlane As CImgPlane, _
    ByVal blueZone As String)

    Call MultiMean(srcPlane, srcZone, "R", redPlane, redZone, "R", idpMultiMeanFuncMean, StdCM_CompRGB2Y, StdCM_CompRGB2Y)
    Call MultiMean(srcPlane, srcZone, "B", bluePlane, blueZone, "B", idpMultiMeanFuncMean, StdCM_CompRGB2Y, StdCM_CompRGB2Y)
    Call MultiMean(srcPlane, srcZone, "GR", greenPlane, greenZone, "GR", idpMultiMeanFuncMean, StdCM_CompRGB2Y, StdCM_CompRGB2Y * 2)

End Sub

Public Sub StdCM_SeparateRGBforClearVid( _
    ByRef srcPlane As CImgPlane, _
    ByVal srcZone As String, _
    ByRef redPlane As CImgPlane, _
    ByVal redZone As String, _
    ByRef greenPlane As CImgPlane, _
    ByVal greenZone As String, _
    ByRef bluePlane As CImgPlane, _
    ByVal blueZone As String)
    
    Dim tmpG3plane As CImgPlane
    Call GetFreePlane(tmpG3plane, "gbayer3", idpDepthS16, False, "tmp3")
    
    Dim tmpG1plane As CImgPlane
    Call GetFreePlane(tmpG1plane, "gBayer", idpDepthS16, False, "tmp1")
    
    Call MultiMean(srcPlane, srcZone, "R", redPlane, redZone, "R", idpMultiMeanFuncMean, StdCM_CompRGB2Y, StdCM_CompRGB2Y)
    Call MultiMean(srcPlane, srcZone, "B", bluePlane, blueZone, "B", idpMultiMeanFuncMean, StdCM_CompRGB2Y, StdCM_CompRGB2Y)
    Call MultiMean(srcPlane, srcZone, "GR", tmpG1plane, greenZone, "GR", idpMultiMeanFuncMean, StdCM_CompRGB2Y, StdCM_CompRGB2Y * 4)
    Call MultiMean(srcPlane, srcZone, "G3", tmpG3plane, "GBAYER3_FULL", "G3", idpMultiMeanFuncMean, StdCM_CompRGB2Y, StdCM_CompRGB2Y * 2)
    Call Add(tmpG1plane, greenZone, idpColorFlat, tmpG3plane, "GBAYER3_FULL", idpColorFlat, _
             greenPlane, greenZone, idpColorFlat)
             
    Call ReleasePlane(tmpG1plane)
    Call ReleasePlane(tmpG3plane)

End Sub

Private Sub stdCM_ApplyLowPassFilterAll( _
    ByRef inPlane As CImgPlane, _
    ByRef workPlane As CImgPlane, _
    ByVal zonePrefix As String, _
    ByVal lowPassFilterH As String, _
    ByVal lowPassFilterV As String)

    Call Copy(inPlane, zonePrefix & "_ZONE2D_EX_LEFT_IN", EEE_COLOR_FLAT, _
              workPlane, zonePrefix & "_ZONE2D_EX_LEFT_OUT", EEE_COLOR_FLAT)
    Call Copy(workPlane, zonePrefix & "_ZONE2D_EX_LEFT_OUT", EEE_COLOR_FLAT, _
              inPlane, zonePrefix & "_ZONE2D_EX_LEFT_OUT", EEE_COLOR_FLAT)
    Call Copy(inPlane, zonePrefix & "_ZONE2D_EX_RIGHT_IN", EEE_COLOR_FLAT, _
              workPlane, zonePrefix & "_ZONE2D_EX_RIGHT_OUT", EEE_COLOR_FLAT)
    Call Copy(workPlane, zonePrefix & "_ZONE2D_EX_RIGHT_OUT", EEE_COLOR_FLAT, _
              inPlane, zonePrefix & "_ZONE2D_EX_RIGHT_OUT", EEE_COLOR_FLAT)
    Call Convolution(inPlane, zonePrefix & "_FULL", EEE_COLOR_FLAT, workPlane, zonePrefix & "_FULL", EEE_COLOR_FLAT, lowPassFilterH)

    Call Copy(workPlane, zonePrefix & "_ZONE2D_EX_TOP_IN", EEE_COLOR_FLAT, _
              inPlane, zonePrefix & "_ZONE2D_EX_TOP_OUT", EEE_COLOR_FLAT)
    Call Copy(inPlane, zonePrefix & "_ZONE2D_EX_TOP_OUT", EEE_COLOR_FLAT, _
              workPlane, zonePrefix & "_ZONE2D_EX_TOP_OUT", EEE_COLOR_FLAT)
    Call Copy(workPlane, zonePrefix & "_ZONE2D_EX_BOTTOM_IN", EEE_COLOR_FLAT, _
              inPlane, zonePrefix & "_ZONE2D_EX_BOTTOM_OUT", EEE_COLOR_FLAT)
    Call Copy(inPlane, zonePrefix & "_ZONE2D_EX_BOTTOM_OUT", EEE_COLOR_FLAT, _
              workPlane, zonePrefix & "_ZONE2D_EX_BOTTOM_OUT", EEE_COLOR_FLAT)
    Call Convolution(workPlane, zonePrefix & "_FULL", EEE_COLOR_FLAT, inPlane, zonePrefix & "_FULL", EEE_COLOR_FLAT, lowPassFilterV)

End Sub

Private Sub stdCMT_ApplyLowPassFilter_L( _
    ByRef inPlane As CImgPlane, _
    ByRef workPlane As CImgPlane, _
    ByVal zonePrefix As String, _
    ByVal lowPassFilterH As String, _
    ByVal lowPassFilterV As String)

    Call Copy(inPlane, zonePrefix & "_ZONE2D_L_EX_LEFT_IN", EEE_COLOR_FLAT, _
              workPlane, zonePrefix & "_ZONE2D_L_EX_LEFT_OUT", EEE_COLOR_FLAT)
    Call Copy(workPlane, zonePrefix & "_ZONE2D_L_EX_LEFT_OUT", EEE_COLOR_FLAT, _
              inPlane, zonePrefix & "_ZONE2D_L_EX_LEFT_OUT", EEE_COLOR_FLAT)
    Call Copy(inPlane, zonePrefix & "_ZONE2D_L_EX_RIGHT_IN", EEE_COLOR_FLAT, _
              workPlane, zonePrefix & "_ZONE2D_L_EX_RIGHT_OUT", EEE_COLOR_FLAT)
    Call Copy(workPlane, zonePrefix & "_ZONE2D_L_EX_RIGHT_OUT", EEE_COLOR_FLAT, _
              inPlane, zonePrefix & "_ZONE2D_L_EX_RIGHT_OUT", EEE_COLOR_FLAT)
    Call Convolution(inPlane, zonePrefix & "_FULL_L", EEE_COLOR_FLAT, workPlane, zonePrefix & "_FULL", EEE_COLOR_FLAT, lowPassFilterH)

    Call Copy(workPlane, zonePrefix & "_ZONE2D_L_EX_TOP_IN", EEE_COLOR_FLAT, _
              inPlane, zonePrefix & "_ZONE2D_L_EX_TOP_OUT", EEE_COLOR_FLAT)
    Call Copy(inPlane, zonePrefix & "_ZONE2D_L_EX_TOP_OUT", EEE_COLOR_FLAT, _
              workPlane, zonePrefix & "_ZONE2D_L_EX_TOP_OUT", EEE_COLOR_FLAT)
    Call Copy(workPlane, zonePrefix & "_ZONE2D_L_EX_BOTTOM_IN", EEE_COLOR_FLAT, _
              inPlane, zonePrefix & "_ZONE2D_L_EX_BOTTOM_OUT", EEE_COLOR_FLAT)
    Call Copy(inPlane, zonePrefix & "_ZONE2D_L_EX_BOTTOM_OUT", EEE_COLOR_FLAT, _
              workPlane, zonePrefix & "_ZONE2D_L_EX_BOTTOM_OUT", EEE_COLOR_FLAT)
    Call Convolution(workPlane, zonePrefix & "_FULL_L", EEE_COLOR_FLAT, inPlane, zonePrefix & "_FULL", EEE_COLOR_FLAT, lowPassFilterV)

End Sub

Private Sub stdCMT_ApplyLowPassFilter_R( _
    ByRef inPlane As CImgPlane, _
    ByRef workPlane As CImgPlane, _
    ByVal zonePrefix As String, _
    ByVal lowPassFilterH As String, _
    ByVal lowPassFilterV As String)

    Call Copy(inPlane, zonePrefix & "_ZONE2D_R_EX_LEFT_IN", EEE_COLOR_FLAT, _
              workPlane, zonePrefix & "_ZONE2D_R_EX_LEFT_OUT", EEE_COLOR_FLAT)
    Call Copy(workPlane, zonePrefix & "_ZONE2D_R_EX_LEFT_OUT", EEE_COLOR_FLAT, _
              inPlane, zonePrefix & "_ZONE2D_R_EX_LEFT_OUT", EEE_COLOR_FLAT)
    Call Copy(inPlane, zonePrefix & "_ZONE2D_R_EX_RIGHT_IN", EEE_COLOR_FLAT, _
              workPlane, zonePrefix & "_ZONE2D_R_EX_RIGHT_OUT", EEE_COLOR_FLAT)
    Call Copy(workPlane, zonePrefix & "_ZONE2D_R_EX_RIGHT_OUT", EEE_COLOR_FLAT, _
              inPlane, zonePrefix & "_ZONE2D_R_EX_RIGHT_OUT", EEE_COLOR_FLAT)
    Call Convolution(inPlane, zonePrefix & "_FULL_R", EEE_COLOR_FLAT, workPlane, zonePrefix & "_FULL", EEE_COLOR_FLAT, lowPassFilterH)

    Call Copy(workPlane, zonePrefix & "_ZONE2D_R_EX_TOP_IN", EEE_COLOR_FLAT, _
              inPlane, zonePrefix & "_ZONE2D_R_EX_TOP_OUT", EEE_COLOR_FLAT)
    Call Copy(inPlane, zonePrefix & "_ZONE2D_R_EX_TOP_OUT", EEE_COLOR_FLAT, _
              workPlane, zonePrefix & "_ZONE2D_R_EX_TOP_OUT", EEE_COLOR_FLAT)
    Call Copy(workPlane, zonePrefix & "_ZONE2D_R_EX_BOTTOM_IN", EEE_COLOR_FLAT, _
              inPlane, zonePrefix & "_ZONE2D_R_EX_BOTTOM_OUT", EEE_COLOR_FLAT)
    Call Copy(inPlane, zonePrefix & "_ZONE2D_R_EX_BOTTOM_OUT", EEE_COLOR_FLAT, _
              workPlane, zonePrefix & "_ZONE2D_R_EX_BOTTOM_OUT", EEE_COLOR_FLAT)
    Call Convolution(workPlane, zonePrefix & "_FULL_R", EEE_COLOR_FLAT, inPlane, zonePrefix & "_FULL", EEE_COLOR_FLAT, lowPassFilterV)

End Sub

Public Sub StdCM_ApplyLowPassFilterYLine( _
    ByRef inPlane As CImgPlane, _
    ByRef workPlane As CImgPlane, _
    ByVal lowPassFilterH As String, _
    ByVal lowPassFilterV As String)

    Call stdCM_ApplyLowPassFilterAll(inPlane, workPlane, "YLINE", lowPassFilterH, lowPassFilterV)
    
End Sub

Public Sub StdCM_ApplyLowPassFilterYLocal( _
    ByRef inPlane As CImgPlane, _
    ByRef workPlane As CImgPlane, _
    ByVal lowPassFilterH As String, _
    ByVal lowPassFilterV As String)

    Call stdCM_ApplyLowPassFilterAll(inPlane, workPlane, "YLOCAL", lowPassFilterH, lowPassFilterV)

End Sub

Public Sub StdCMT_ApplyLowPassFilterYLine( _
    ByRef inPlane As CImgPlane, _
    ByRef workPlane As CImgPlane, _
    ByVal lowPassFilterH As String, _
    ByVal lowPassFilterV As String)

    Call stdCMT_ApplyLowPassFilter_L(inPlane, workPlane, "YLINE", lowPassFilterH, lowPassFilterV)
    Call stdCMT_ApplyLowPassFilter_R(inPlane, workPlane, "YLINE", lowPassFilterH, lowPassFilterV)
    
End Sub

Public Sub StdCM_MakeRyBy( _
    ByRef redPlane As CImgPlane, _
    ByRef greenPlane As CImgPlane, _
    ByRef bluePlane As CImgPlane, _
    ByRef ryPlane As CImgPlane, _
    ByRef byPlane As CImgPlane)
    
    Dim rMean() As Double
    Dim gMean() As Double
    Dim bMean() As Double
    TheResult.GetResult "R_MEAN", rMean
    TheResult.GetResult "G_MEAN", gMean
    TheResult.GetResult "B_MEAN", bMean

    Dim RyR_var(nSite) As Double
    Dim RyG_var(nSite) As Double
    Dim RyB_var(nSite) As Double
    Dim ByR_var(nSite) As Double
    Dim ByG_var(nSite) As Double
    Dim ByB_var(nSite) As Double
    
    Dim site As Long
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            RyG_var(site) = RyGain * (0.59 + 0.59 * ByHue) * MagicNumber
            RyB_var(site) = RyGain * (0.11 - 0.89 * ByHue) * MagicNumber
            RyR_var(site) = RyG_var(site) + RyB_var(site)
            RyR_var(site) = stdCM_div(RyR_var(site), rMean(site))
            RyG_var(site) = stdCM_div(RyG_var(site), gMean(site))
            RyB_var(site) = stdCM_div(RyB_var(site), bMean(site))
            
            ByR_var(site) = ByGain * (0.3 - 0.7 * RyHue) * MagicNumber
            ByG_var(site) = ByGain * (0.59 + 0.59 * RyHue) * MagicNumber
            ByB_var(site) = ByR_var(site) + ByG_var(site)
            ByR_var(site) = stdCM_div(ByR_var(site), rMean(site))
            ByG_var(site) = stdCM_div(ByG_var(site), gMean(site))
            ByB_var(site) = stdCM_div(ByB_var(site), bMean(site))
        End If
    Next site
    
    Dim workPlane1 As CImgPlane
    Dim workPlane2 As CImgPlane
    Call GetFreePlane(workPlane1, redPlane.planeGroup, idpDepthF32, False, "work plane1")
    Call GetFreePlane(workPlane2, redPlane.planeGroup, idpDepthF32, False, "work plane2")

    Call MultiplyConst(redPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, RyR_var, workPlane1, "YLOCAL_ZONE2D", EEE_COLOR_FLAT)
    Call MultiplyConst(greenPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, RyG_var, ryPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT)
    Call Subtract(workPlane1, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, ryPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, workPlane2, "YLOCAL_ZONE2D", EEE_COLOR_FLAT)
    Call MultiplyConst(bluePlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, RyB_var, workPlane1, "YLOCAL_ZONE2D", EEE_COLOR_FLAT)
    Call Subtract(workPlane2, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, workPlane1, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, ryPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT)
    
    Call MultiplyConst(bluePlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, ByB_var, workPlane1, "YLOCAL_ZONE2D", EEE_COLOR_FLAT)
    Call MultiplyConst(greenPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, ByG_var, byPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT)
    Call Subtract(workPlane1, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, byPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, workPlane2, "YLOCAL_ZONE2D", EEE_COLOR_FLAT)
    Call MultiplyConst(redPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, ByR_var, workPlane1, "YLOCAL_ZONE2D", EEE_COLOR_FLAT)
    Call Subtract(workPlane2, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, workPlane1, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, byPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT)

    Call ReleasePlane(workPlane1)
    Call ReleasePlane(workPlane2)
    
End Sub


Private Function stdCM_div( _
    ByVal val1 As Double, _
    ByVal val2 As Double, _
    Optional ByVal errVal As Double = 0) As Double
    If val2 <> 0# Then
        stdCM_div = val1 / val2
    Else
        stdCM_div = errVal
    End If
End Function

Public Sub StdCM_DegcalNew( _
    ByRef ryPlane As CImgPlane, _
    ByRef byPlane As CImgPlane, _
    ByRef degPlane As CImgPlane, _
    ByVal dstZone As String)
    
    Const IntTangentMultiplyer As Double = 128

    'DEG=tan(Ry/By)において、By=0の場合はRyの符号に応じて処理が必要になりますので、フラグを保持しておきます。
    Call PutFlag(ryPlane, dstZone, EEE_COLOR_FLAT, idpCountAbove, 0, 0, idpLimitExclude, "FLG_RY-POSI")
    Call PutFlag(ryPlane, dstZone, EEE_COLOR_FLAT, idpCountBelow, 0, 0, idpLimitExclude, "FLG_RY-NEGA")
    Call PutFlag(byPlane, dstZone, EEE_COLOR_FLAT, idpCountBetween, 0, 0, idpLimitInclude, "FLG_BY-ZERO")
    Call PutFlag(byPlane, dstZone, EEE_COLOR_FLAT, idpCountBelow, 0, 0, idpLimitExclude, "FLG_BY-NEGA")
    
    'By=0, Ry>0なら90度、By=0, Ry<0なら270度となりますので、それぞれのフラグを保持します。尚、90度、270度に
    '関しては後で別理由のフラグと合成する必要がありますので、"TEMP"としておきます。
    Call SharedFlagAnd(degPlane.planeGroup, dstZone, "FLG_090DEG", "FLG_BY-ZERO", "FLG_RY-POSI")
    Call SharedFlagAnd(degPlane.planeGroup, dstZone, "FLG_270DEG", "FLG_BY-ZERO", "FLG_RY-NEGA")
    Call degPlane.GetSharedFlagPlane("FLG_RY-POSI").RemoveFlagBit("FLG_RY-POSI")
    Call degPlane.GetSharedFlagPlane("FLG_RY-NEGA").RemoveFlagBit("FLG_RY-NEGA")
    Call degPlane.GetSharedFlagPlane("FLG_BY-ZERO").RemoveFlagBit("FLG_BY-ZERO")

    'tan(ry/by)を求めます。
    Dim tanPlane As CImgPlane
    Call GetFreePlane(tanPlane, ryPlane.planeGroup, idpDepthF32, False, "tangent plane")
    Call Divide(ryPlane, dstZone, EEE_COLOR_FLAT, byPlane, dstZone, EEE_COLOR_FLAT, tanPlane, dstZone, EEE_COLOR_FLAT)
    
    '128(上記定数)をかけて、整数プレーンにコピーします。
    Dim fWorkPlane1 As CImgPlane
    Call GetFreePlane(fWorkPlane1, ryPlane.planeGroup, idpDepthF32, False, "float work plane1")
    Call MultiplyConst(tanPlane, dstZone, EEE_COLOR_FLAT, IntTangentMultiplyer, fWorkPlane1, dstZone, EEE_COLOR_FLAT)
    Call PutFlag(fWorkPlane1, dstZone, EEE_COLOR_FLAT, idpCountAbove, 2 ^ 15 - 2, 2 ^ 15 - 2, idpLimitInclude, "FLG_TAN_OVER_16BIT")
    Call PutFlag(fWorkPlane1, dstZone, EEE_COLOR_FLAT, idpCountBelow, -2 ^ 15 + 2, -2 ^ 15 + 2, idpLimitInclude, "FLG_TAN_UNDER_16BIT")
    
    Dim iWorkPlane1 As CImgPlane
    Call GetFreePlane(iWorkPlane1, ryPlane.planeGroup, idpDepthS16, False, "integer work plane1")
    Call Copy(fWorkPlane1, dstZone, EEE_COLOR_FLAT, iWorkPlane1, dstZone, EEE_COLOR_FLAT)
    Call iWorkPlane1.WritePixel(2 ^ 15 - 2, EEE_COLOR_FLAT, , "FLG_TAN_OVER_16BIT")
    Call iWorkPlane1.WritePixel(-2 ^ 15 + 2, EEE_COLOR_FLAT, , "FLG_TAN_UNDER_16BIT")
    Call ReleasePlane(tanPlane)
    Call ReleasePlane(fWorkPlane1)
    
    '整数プレーンにコピーされたものを用いて、LUTでarctanを求めます。
    Call ExecuteLUT(iWorkPlane1, dstZone, EEE_COLOR_FLAT, degPlane, dstZone, EEE_COLOR_FLAT, LookUpTableArctan)
    Call ReleasePlane(iWorkPlane1)
    
    '第二・第三象限については180を足しておきます。
    '入力と出力のメモリが同一になるが、コピーしてわざわざ別にするよりもこの場合はこの方が速い。
    Call degPlane.Add(degPlane, 180, EEE_COLOR_FLAT, EEE_COLOR_FLAT, , "FLG_BY-NEGA")
    Call degPlane.GetSharedFlagPlane("FLG_BY-NEGA").RemoveFlagBit("FLG_BY-NEGA")
    
    'By=0(分母0)の点だけケアします。
    Call degPlane.WritePixel(90, EEE_COLOR_FLAT, , "FLG_090DEG")
    Call degPlane.WritePixel(270, EEE_COLOR_FLAT, , "FLG_270DEG")
    Call degPlane.GetSharedFlagPlane("FLG_090DEG").RemoveFlagBit("FLG_090DEG")
    Call degPlane.GetSharedFlagPlane("FLG_270DEG").RemoveFlagBit("FLG_270DEG")
    Call degPlane.GetSharedFlagPlane("FLG_TAN_OVER_16BIT").RemoveFlagBit("FLG_TAN_OVER_16BIT")
    Call degPlane.GetSharedFlagPlane("FLG_TAN_UNDER_16BIT").RemoveFlagBit("FLG_TAN_UNDER_16BIT")
    
End Sub

Public Sub std_CalcYline( _
     ByRef yPlane As CImgPlane, ByVal pSlice As Double, ByVal pDif As Double, ByVal pCoef As Double, ByRef rtnYline() As Double)

    Dim site As Long

    Dim HiLimit(nSite) As Double, LoLimit(nSite) As Double
    Dim yLevel() As Double
    TheResult.GetResult "Y_MEAN", yLevel
    '///// Y MURA SLICE LEVEL /////////////////////////////
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            HiLimit(site) = yLevel(site) * pSlice * pCoef
            LoLimit(site) = HiLimit(site) * (-1)
        End If
    Next site

    Dim hDifPlane As CImgPlane, vDifPlane As CImgPlane
    Call GetFreePlane(hDifPlane, yPlane.planeGroup, idpDepthF32, , "hDifPlane")
    Call GetFreePlane(vDifPlane, yPlane.planeGroup, idpDepthF32, , "vDifPlane")

    '========== H/V LINE DIFF. ============================
    Dim workPlane1 As CImgPlane
    Call GetFreePlane(workPlane1, yPlane.planeGroup, idpDepthF32)
    Call Copy(yPlane, "YLINE_COL_SUB_TARGET", EEE_COLOR_FLAT, workPlane1, "YLINE_COL_SUB_TARGET", EEE_COLOR_FLAT)
    Call Subtract(yPlane, "YLINE_COL_SUB_SOURCE", EEE_COLOR_FLAT, _
                  workPlane1, "YLINE_COL_SUB_TARGET", EEE_COLOR_FLAT, _
                  hDifPlane, "YLINE_COL_SUB_SOURCE", EEE_COLOR_FLAT)
    Call ReleasePlane(workPlane1)

    Call SubRows(yPlane, "YLINE_ROW_SUB", EEE_COLOR_FLAT, vDifPlane, "YLINE_ROW_SUB", EEE_COLOR_FLAT, pDif)

    Dim mYlineH(nSite) As Double, mYlineV(nSite) As Double, mYline(nSite) As Double
    '========== YLINE COUNT ===============================
    Call Count(hDifPlane, "YLINE_COL_JUDGE", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mYlineH, "FLG_YLINE_H")
    Call Count(vDifPlane, "YLINE_ROW_JUDGE", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mYlineV, "FLG_YLINE_V")
    '[IDV Point] "hDifPlane"=H Diff Plane (Flag="FLG_YLINE_H")
    '[IDV Point] "vDifPlane"=V Diff Plane (Flag="FLG_YLINE_V")
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            mYline(site) = mYlineH(site) + mYlineV(site)
        End If
    Next site
    
    Call ReleasePlane(hDifPlane)
    Call ReleasePlane(vDifPlane)

    '========== SKIP CHECK ================================
    Dim IsSkip As Boolean
    IsSkip = True
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If mYline(site) > 0 Then IsSkip = False
        End If
    Next site
    If IsSkip = True Then
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                rtnYline(site) = 0
            End If
        Next site
        GoTo EndYLine
    End If

    Dim FlgPlane As CImgPlane
    Call GetFreePlane(FlgPlane, yPlane.planeGroup, idpDepthS16, True, "flgPlane")
    '========== MAKE YLINE DEFFECT IMAGE ==================
    Call SharedFlagOr(FlgPlane.planeGroup, "YLINE_ZONE2D", "FLG_YLINE_HV", "FLG_YLINE_H", "FLG_YLINE_V")
    Call FlagCopy(FlgPlane, "YLINE_ZONE2D", "FLG_YLINE_HV")
    Call FlgPlane.GetSharedFlagPlane("FLG_YLINE_HV").RemoveFlagBit("FLG_YLINE_HV")
    Call FlgPlane.GetSharedFlagPlane("FLG_YLINE_H").RemoveFlagBit("FLG_YLINE_H")
    Call FlgPlane.GetSharedFlagPlane("FLG_YLINE_V").RemoveFlagBit("FLG_YLINE_V")

    Dim tmpPlane As CImgPlane
    Call GetFreePlane(tmpPlane, yPlane.planeGroup, idpDepthS16, True, "tmpPlane")
    '========== 5BIT MEDIAN ===============================
    Call Median(FlgPlane, "YLINE_FULL", EEE_COLOR_FLAT, tmpPlane, "YLINE_FULL", EEE_COLOR_FLAT, 5, 1)
    Call Median(tmpPlane, "YLINE_FULL", EEE_COLOR_FLAT, FlgPlane, "YLINE_FULL", EEE_COLOR_FLAT, 1, 5)
    Call ReleasePlane(tmpPlane)

    Dim compPlane1 As CImgPlane
    Call GetFreePlane(compPlane1, "pylinecomp", idpDepthS16, False, "compPlane")
    '========== MULTIMEAN =================================
    Call MultiMean(FlgPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, compPlane1, "YLINECOMP_FULL", EEE_COLOR_FLAT, idpMultiMeanFuncSum, yLineCompPix, yLineCompPix)
    Call ReleasePlane(FlgPlane)

    Dim rtnVal(nSite) As Double
    Dim compPlane2 As CImgPlane
    Call GetFreePlane(compPlane2, compPlane1.planeGroup, idpDepthS16, False, "comp convolution plane")
    Call CountForFlgBitImgPlane(compPlane1, "YLINECOMP_FULL", EEE_COLOR_FLAT, idpCountAbove, 1, 1, idpLimitExclude, rtnVal, compPlane2, 1)

    Call Convolution(compPlane2, "YLINECOMP_FULL", EEE_COLOR_FLAT, compPlane1, "YLINECOMP_FULL", EEE_COLOR_FLAT, Kernel_YLine)

    '========== COUNT YLINE ===============================
    Call Count(compPlane1, "YLINECOMP_FULL", EEE_COLOR_FLAT, idpCountAbove, yLineCount, yLineCount, idpLimitExclude, rtnYline)
    '[IDV Point] "compPlane"=YLINE Final Judgement Plane
    Call ReleasePlane(compPlane1)
    Call ReleasePlane(compPlane2)

EndYLine:
    Call yPlane.GetSharedFlagPlane("FLG_YLINE_HV").RemoveFlagBit("FLG_YLINE_HV")
    Call yPlane.GetSharedFlagPlane("FLG_YLINE_H").RemoveFlagBit("FLG_YLINE_H")
    Call yPlane.GetSharedFlagPlane("FLG_YLINE_V").RemoveFlagBit("FLG_YLINE_V")

End Sub

Private Sub std_CalcYglob2(ByRef yPlane As CImgPlane, ByVal pSlice As Double, ByVal pHdif As Double, ByVal pVdif As Double, ByRef rtnYglob2() As Double, ByRef rtnYgUp2() As Double)

    Dim site As Long
    
    Dim HiLimit(nSite) As Double, LoLimit(nSite) As Double
    Dim yLevel() As Double
    TheResult.GetResult "Y_MEAN", yLevel
    '///// Y MURA SLICE LEVEL ////////////////////
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            HiLimit(site) = (yLevel(site) * pSlice) ^ 2
            LoLimit(site) = HiLimit(site) * (-1)
        End If
    Next site

    Dim hDifPlane As CImgPlane, vDifPlane As CImgPlane
    Call GetFreePlane(hDifPlane, yPlane.planeGroup, idpDepthF32, , "hDifPlane")
    Call GetFreePlane(vDifPlane, yPlane.planeGroup, idpDepthF32, , "vDifPlane")
    Dim tmp1Plane As CImgPlane, tmp2Plane As CImgPlane, tmp3Plane As CImgPlane
    Call GetFreePlane(tmp1Plane, yPlane.planeGroup, idpDepthF32, , "tmp1Plane")
    Call GetFreePlane(tmp2Plane, yPlane.planeGroup, idpDepthF32, , "tmp2Plane")
    Call GetFreePlane(tmp3Plane, yPlane.planeGroup, idpDepthF32, , "tmp3Plane")

    Call Copy(yPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, tmp1Plane, "YLINE_ZONE2D", EEE_COLOR_FLAT)

    '/*---------------------- H DIF ---------------------*/
    Call Subtract(yPlane, "YLINE_ZONE2D_YG2_H", EEE_COLOR_FLAT, tmp1Plane, "YLINE_ZONE2D_YG2_LSFT", EEE_COLOR_FLAT, tmp2Plane, "YLINE_ZONE2D_YG2_H", EEE_COLOR_FLAT)
    Call Subtract(yPlane, "YLINE_ZONE2D_YG2_H", EEE_COLOR_FLAT, tmp1Plane, "YLINE_ZONE2D_YG2_RSFT", EEE_COLOR_FLAT, tmp3Plane, "YLINE_ZONE2D_YG2_H", EEE_COLOR_FLAT)
    Call Add(tmp2Plane, "YLINE_ZONE2D_YG2_H", EEE_COLOR_FLAT, tmp3Plane, "YLINE_ZONE2D_YG2_H", EEE_COLOR_FLAT, hDifPlane, "YLINE_ZONE2D_YG2_H", EEE_COLOR_FLAT)
    Call DivideConst(hDifPlane, "YLINE_ZONE2D_YG2_H", EEE_COLOR_FLAT, 2, hDifPlane, "YLINE_ZONE2D_YG2_H", EEE_COLOR_FLAT)

    '/*---------------------- V DIF ---------------------*/
    Call Subtract(yPlane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT, tmp1Plane, "YLINE_ZONE2D_YG2_TSFT", EEE_COLOR_FLAT, tmp2Plane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT)
    Call Subtract(yPlane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT, tmp1Plane, "YLINE_ZONE2D_YG2_BSFT", EEE_COLOR_FLAT, tmp3Plane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT)
    Call Add(tmp2Plane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT, tmp3Plane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT, vDifPlane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT)
    Call DivideConst(vDifPlane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT, 2, vDifPlane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT)

    Call ReleasePlane(tmp1Plane)
    Call ReleasePlane(tmp2Plane)
    Call ReleasePlane(tmp3Plane)

    Dim hvDifPlane As CImgPlane
    Call GetFreePlane(hvDifPlane, yPlane.planeGroup, idpDepthF32, True, "hvdifPlane")
    'X^2
    Call Multiply(hDifPlane, "YLINE_ZONE2D_YG2_H", EEE_COLOR_FLAT, hDifPlane, "YLINE_ZONE2D_YG2_H", EEE_COLOR_FLAT, hDifPlane, "YLINE_ZONE2D_YG2_H", EEE_COLOR_FLAT)
    'Y^2
    Call Multiply(vDifPlane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT, vDifPlane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT, vDifPlane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT)
    'X^2+Y^2
    Call Add(hDifPlane, "YLINE_ZONE2D_YG2_HV", EEE_COLOR_FLAT, vDifPlane, "YLINE_ZONE2D_YG2_HV", EEE_COLOR_FLAT, hvDifPlane, "YLINE_ZONE2D_YG2_HV", EEE_COLOR_FLAT)

    Call PutFlag(hvDifPlane, "YLINE_ZONE2D_YG2_HV", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, "FLG_YGLOB2")

    Call ReleasePlane(hDifPlane)
    Call ReleasePlane(vDifPlane)

    Dim FlgPlane As CImgPlane, convPlane As CImgPlane
    Call GetFreePlane(FlgPlane, yPlane.planeGroup, idpDepthS16, True, "flgPlane")
    Call GetFreePlane(convPlane, yPlane.planeGroup, idpDepthS16, True, "convPlane")
    '========== CONTINUATION POINT DEFECT =================
    Call FlagCopy(FlgPlane, "YLINE_ZONE2D_YG2_HV", "FLG_YGLOB2")
    Call Convolution(FlgPlane, "YLINE_ZONE2D_YG2_HV", EEE_COLOR_FLAT, convPlane, "YLINE_ZONE2D_YG2_HV", EEE_COLOR_FLAT, "Kernel_YGlobal2")
    Call Count(convPlane, "YLINE_ZONE2D_YG2_HV", EEE_COLOR_FLAT, idpCountAbove, yGlob2Count, yGlob2Count, idpLimitExclude, rtnYglob2, "FLG_YGUP2")

    Dim tmpYgUp2(nSite) As Double
    Dim tmpMin(nSite) As Double, tmpMax(nSite) As Double
    Call MinMax(hvDifPlane, "YLINE_ZONE2D_YG2_HV", EEE_COLOR_FLAT, tmpMin, tmpMax, "FLG_YGUP2")
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            If rtnYglob2(site) > 0 And yLevel(site) > 0 Then
                If Abs(tmpMax(site)) > Abs(tmpMin(site)) Then
                    tmpYgUp2(site) = tmpMax(site)
                Else
                    tmpYgUp2(site) = tmpMin(site)
                End If
                rtnYgUp2(site) = Div(Sqr(Abs(tmpYgUp2(site))), yLevel(site), 999)
            Else
                rtnYgUp2(site) = 0
            End If
        End If
    Next site

    Call yPlane.GetSharedFlagPlane("FLG_YGLOB2").RemoveFlagBit("FLG_YGLOB2")
    Call yPlane.GetSharedFlagPlane("FLG_YGUP2").RemoveFlagBit("FLG_YGUP2")
End Sub

Public Sub std_CalcYlocal( _
    ByRef hDifPlane As CImgPlane, ByRef vDifPlane As CImgPlane, ByVal pSlice As Double, ByVal pDif As Double, ByVal pCoef As Double, ByRef rtnYlocal() As Double, ByRef rtnYlUpp() As Double)

    Dim site As Long

    Dim HiLimit(nSite) As Double, LoLimit(nSite) As Double
    Dim yLevel() As Double
    TheResult.GetResult "Y_MEAN", yLevel
    '///// Y MURA SLICE LEVEL ////////////////////
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            HiLimit(site) = yLevel(site) * pSlice * pCoef
            LoLimit(site) = HiLimit(site) * (-1)
        End If
    Next site

    Dim localPlane As CImgPlane
    Call GetFreePlane(localPlane, "pylocal", idpDepthF32, False, "local source")
    
    Dim mYlocal(nSite) As Double, mYlocalH(nSite) As Double, mYlocalV(nSite) As Double
    '========== YLOCAL COUNT ==============================
    Call Count(hDifPlane, "YLOCAL_COL_JUDGE", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mYlocalH, "FLG_YLOCAL_H")
    Call Count(vDifPlane, "YLOCAL_ROW_JUDGE", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mYlocalV, "FLG_YLOCAL_V")
    '[IDV Point] "hDifPlane"=H Diff Plane (Flag="FLG_YLOCAL_H")
    '[IDV Point] "vDifPlane"=V Diff Plane (Flag="FLG_YLOCAL_V")
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            mYlocal(site) = mYlocalH(site) + mYlocalV(site)
        End If
    Next site
    
    '========== SKIP CHECK ================================
    Dim IsSkip As Boolean
    IsSkip = True
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If mYlocal(site) > 0 Then IsSkip = False
        End If
    Next site
    If IsSkip = True Then
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                rtnYlocal(site) = 0
                rtnYlUpp(site) = 0
            End If
        Next site
        GoTo EndYlocal
    End If

    Dim FlgPlane As CImgPlane
    Call GetFreePlane(FlgPlane, localPlane.planeGroup, idpDepthS16, True, "Flag Plane")
    Call SharedFlagOr(FlgPlane.planeGroup, "YLOCAL_ZONE2D", "FLG_YLOCAL_HV", "FLG_YLOCAL_H", "FLG_YLOCAL_V")
    Call FlagCopy(FlgPlane, "YLOCAL_ZONE2D", "FLG_YLOCAL_HV")
    Call localPlane.GetSharedFlagPlane("FLG_YLOCAL_HV").RemoveFlagBit("FLG_YLOCAL_HV")
    Call localPlane.GetSharedFlagPlane("FLG_YLOCAL_H").RemoveFlagBit("FLG_YLOCAL_H")
    Call localPlane.GetSharedFlagPlane("FLG_YLOCAL_V").RemoveFlagBit("FLG_YLOCAL_V")

    Dim cfuPlane As CImgPlane
    Call GetFreePlane(cfuPlane, localPlane.planeGroup, idpDepthS16, , "cfuPlane")
    Call Convolution(FlgPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, cfuPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, Kernel_Ylocal)
    Call ReleasePlane(FlgPlane)

    '========== COUNT YLOCAL ==============================
    Call Count(cfuPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, idpCountAbove, yLocalCount, yLocalCount, idpLimitExclude, rtnYlocal, "FLG_YLOCAL")
    '[IDV Point] "cfuPlane"=YLOCAL Final Judgement Plane
    Call ReleasePlane(cfuPlane)


    Dim h_temp(nSite) As Double, v_temp(nSite) As Double
    Call AbsMax(hDifPlane, "YLOCAL_COL_JUDGE", EEE_COLOR_FLAT, h_temp, "FLG_YLOCAL")
    Call AbsMax(vDifPlane, "YLOCAL_ROW_JUDGE", EEE_COLOR_FLAT, v_temp, "FLG_YLOCAL")
    '[IDV Point] "hDifPlane"=YL_UPP Judgement Plane (H Diff)(InputFlag="FLG_YLOCAL")
    '[IDV Point] "vDifPlane"=YL_UPP Judgement Plane (V Diff)(InputFlag="FLG_YLOCAL")
    Call localPlane.GetSharedFlagPlane("FLG_YLOCAL").RemoveFlagBit("FLG_YLOCAL")
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If Abs(h_temp(site)) = 1E+38 Then h_temp(site) = 0
            If Abs(v_temp(site)) = 1E+38 Then v_temp(site) = 0
            rtnYlUpp(site) = StdCM_GetMax(Abs(h_temp(site)), Abs(v_temp(site))) / pCoef
        End If
    Next site

   
EndYlocal:
    Call localPlane.GetSharedFlagPlane("FLG_YLOCAL_HV").RemoveFlagBit("FLG_YLOCAL_HV")
    Call localPlane.GetSharedFlagPlane("FLG_YLOCAL_H").RemoveFlagBit("FLG_YLOCAL_H")
    Call localPlane.GetSharedFlagPlane("FLG_YLOCAL_V").RemoveFlagBit("FLG_YLOCAL_V")
    Call localPlane.GetSharedFlagPlane("FLG_YLOCAL").RemoveFlagBit("FLG_YLOCAL")

End Sub
Public Sub std_CalcYfrm2d( _
    ByRef hDifPlane As CImgPlane, ByRef vDifPlane As CImgPlane, ByVal pSlice As Double, ByVal pDif As Double, ByVal pCoef As Double, ByRef rtnYfrm2d() As Double, ByRef rtnYfupd() As Double)

    Dim site As Long

    Dim HiLimit(nSite) As Double, LoLimit(nSite) As Double
    Dim yLevel() As Double
    TheResult.GetResult "Y_MEAN", yLevel
    '///// Y MURA SLICE LEVEL ////////////////////
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            HiLimit(site) = yLevel(site) * pSlice * pCoef
            LoLimit(site) = HiLimit(site) * (-1)
        End If
    Next site

    Dim mYfrm2d(nSite) As Double, mYfrm2dH(nSite) As Double, mYfrm2dV(nSite) As Double
    
    '========== YFRM2D COUNT ==============================
    Call Count(hDifPlane, "YLOCAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mYfrm2dH, "FLG_YFRAME2D_H")
    Call Count(vDifPlane, "YLOCAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mYfrm2dV, "FLG_YFRAME2D_V")
    '[IDV Point] "hDifPlane"=H Diff Plane (Flag="FLG_YFRAME2D_H")
    '[IDV Point] "vDifPlane"=V Diff Plane (Flag="FLG_YFRAME2D_V")
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            mYfrm2d(site) = mYfrm2dH(site) + mYfrm2dV(site)
        End If
    Next site
    
    '========== SKIP CHECK ================================
    Dim IsSkip As Boolean
    IsSkip = True
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If mYfrm2d(site) > 0 Then IsSkip = False
        End If
    Next site
    If IsSkip = True Then
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                rtnYfrm2d(site) = 0
                rtnYfupd(site) = 0
            End If
        Next site
        GoTo EndYfrm2d
    End If

    Dim FlgPlane As CImgPlane
    Call GetFreePlane(FlgPlane, hDifPlane.planeGroup, idpDepthS16, True, "flgPlane")
    '========== MAKE Y FRM2D DEFFECT IMAGE ================
    Call SharedFlagOr(FlgPlane.planeGroup, "YLOCAL_ZONE2D", "FLG_YFRAME2D_HV", "FLG_YFRAME2D_H", "FLG_YFRAME2D_V")
    Call FlagCopy(FlgPlane, "YLOCAL_ZONE2D", "FLG_YFRAME2D_HV")
    Call FlgPlane.GetSharedFlagPlane("FLG_YFRAME2D_HV").RemoveFlagBit("FLG_YFRAME2D_HV")
    Call FlgPlane.GetSharedFlagPlane("FLG_YFRAME2D_H").RemoveFlagBit("FLG_YFRAME2D_H")
    Call FlgPlane.GetSharedFlagPlane("FLG_YFRAME2D_V").RemoveFlagBit("FLG_YFRAME2D_V")

    Dim cfuPlane As CImgPlane
    Call GetFreePlane(cfuPlane, FlgPlane.planeGroup, idpDepthS16, False, "cfuPlane")
    '========== CONVOLUTION ===============================
    Call Convolution(FlgPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, cfuPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, Kernel_YFrame2D)
    Call ReleasePlane(FlgPlane)

    '========== ZONE2 DATA CLEAR ==========================
    Call WritePixel(cfuPlane, "YLOCAL_ZONE2", EEE_COLOR_FLAT, 0)

    '========== COUNT YFRM2D ==============================
    Call Count(cfuPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, idpCountAbove, yLocalCount, yLocalCount, idpLimitExclude, rtnYfrm2d, "FLG_YFRAME2D")
    '[IDV Point] "cfuPlane"=YFRM2D Final Judgement Plane
    Call ReleasePlane(cfuPlane)

    Dim h_temp(nSite) As Double, v_temp(nSite) As Double
    Call AbsMax(hDifPlane, "YLOCAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, h_temp, "FLG_YFRAME2D")
    Call AbsMax(vDifPlane, "YLOCAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, v_temp, "FLG_YFRAME2D")
    '[IDV Point] "hDifPlane"=YF_UPD Judgement Plane (H Diff)(InputFlag="FLG_YFRAME2D")
    '[IDV Point] "vDifPlane"=YF_UPD Judgement Plane (V Diff)(InputFlag="FLG_YFRAME2D")
    Call hDifPlane.GetSharedFlagPlane("FLG_YFRAME2D").RemoveFlagBit("FLG_YFRAME2D")
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If Abs(h_temp(site)) = 1E+38 Then h_temp(site) = 0
            If Abs(v_temp(site)) = 1E+38 Then v_temp(site) = 0
            rtnYfupd(site) = StdCM_GetMax(Abs(h_temp(site)), Abs(v_temp(site))) / pCoef
        End If
    Next site


EndYfrm2d:
    Call hDifPlane.GetSharedFlagPlane("FLG_YFRAME2D_HV").RemoveFlagBit("FLG_YFRAME2D_HV")
    Call hDifPlane.GetSharedFlagPlane("FLG_YFRAME2D_H").RemoveFlagBit("FLG_YFRAME2D_H")
    Call hDifPlane.GetSharedFlagPlane("FLG_YFRAME2D_V").RemoveFlagBit("FLG_YFRAME2D_V")
    Call hDifPlane.GetSharedFlagPlane("FLG_YFRAME2D").RemoveFlagBit("FLG_YFRAME2D")

End Sub

Public Sub std_CalcYfrm2( _
    ByRef hDifPlane As CImgPlane, ByRef vDifPlane As CImgPlane, ByVal pSlice As Double, ByVal pDif As Double, ByVal pCoef As Double, ByRef rtnYfrm2() As Double, ByRef rtnYfup2() As Double)

    Dim site As Long

    Dim HiLimit(nSite) As Double, LoLimit(nSite) As Double
    Dim yLevel() As Double
    TheResult.GetResult "Y_MEAN", yLevel
    '///// Y MURA SLICE LEVEL ////////////////////
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            HiLimit(site) = yLevel(site) * pSlice * pCoef
            LoLimit(site) = HiLimit(site) * (-1)
        End If
    Next site

    Dim mYfrm2(nSite) As Double, mYfrm2H(nSite) As Double, mYfrm2V(nSite) As Double
    '========== YFRM2 COUNT ===============================
    Call Count(hDifPlane, "YLOCAL_ZONE2_CFU", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mYfrm2H, "FLG_YFRAME2_H")
    Call Count(vDifPlane, "YLOCAL_ZONE2_CFU", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mYfrm2V, "FLG_YFRAME2_V")
    '[IDV Point] "hDifPlane"=H Diff Plane (Flag="FLG_YFRAME2_H")
    '[IDV Point] "vDifPlane"=V Diff Plane (Flag="FLG_YFRAME2_V")
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            mYfrm2(site) = mYfrm2H(site) + mYfrm2V(site)
        End If
    Next site
    
    '========== SKIP CHECK ================================
    Dim IsSkip As Boolean
    IsSkip = True
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If mYfrm2(site) > 0 Then IsSkip = False
        End If
    Next site
    If IsSkip = True Then
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                rtnYfrm2(site) = 0
                rtnYfup2(site) = 0
            End If
        Next site
        GoTo EndYfrm2
    End If

    Dim FlgPlane As CImgPlane
    Call GetFreePlane(FlgPlane, hDifPlane.planeGroup, idpDepthS16, True, "flgPlane")
    '========== MAKE YFRM2 DEFFECT IMAGE ==================
    Call SharedFlagOr(FlgPlane.planeGroup, "YLOCAL_ZONE2_CFU", "FLG_YFRAME2_HV", "FLG_YFRAME2_H", "FLG_YFRAME2_V")
    Call FlagCopy(FlgPlane, "YLOCAL_ZONE2_CFU", "FLG_YFRAME2_HV")
    Call FlgPlane.GetSharedFlagPlane("FLG_YFRAME2_HV").RemoveFlagBit("FLG_YFRAME2_HV")
    Call FlgPlane.GetSharedFlagPlane("FLG_YFRAME2_H").RemoveFlagBit("FLG_YFRAME2_H")
    Call FlgPlane.GetSharedFlagPlane("FLG_YFRAME2_V").RemoveFlagBit("FLG_YFRAME2_V")

    Dim cfuPlane As CImgPlane
    Call GetFreePlane(cfuPlane, hDifPlane.planeGroup, idpDepthS16, True, "cfuPlane")
    '========== CONVOLUTION ===============================
    Call Convolution(FlgPlane, "YLOCAL_ZONE2_CFU", EEE_COLOR_FLAT, cfuPlane, "YLOCAL_ZONE2_CFU", EEE_COLOR_FLAT, Kernel_YFrame2)
    Call ReleasePlane(FlgPlane)

    '========== COUNT YFRM2 ===============================
    Call Count(cfuPlane, "YLOCAL_ZONE2", EEE_COLOR_FLAT, idpCountAbove, yLocalCount, yLocalCount, idpLimitExclude, rtnYfrm2, "FLG_YFRAME2")
    '[IDV Point] "cfuPlane"=YFRM2 Final Judgement Plane
    Call ReleasePlane(cfuPlane)

    Dim h_temp(nSite) As Double, v_temp(nSite) As Double
    Call AbsMax(hDifPlane, "YLOCAL_ZONE2", EEE_COLOR_FLAT, h_temp, "FLG_YFRAME2")
    Call AbsMax(vDifPlane, "YLOCAL_ZONE2", EEE_COLOR_FLAT, v_temp, "FLG_YFRAME2")
    '[IDV Point] "hDifPlane"=YF_UP2 Judgement Plane (H Diff)(InputFlag="FLG_YFRAME2")
    '[IDV Point] "vDifPlane"=YF_UP2 Judgement Plane (V Diff)(InputFlag="FLG_YFRAME2")
    Call hDifPlane.GetSharedFlagPlane("FLG_YFRAME2").RemoveFlagBit("FLG_YFRAME2")
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If Abs(h_temp(site)) = 1E+38 Then h_temp(site) = 0
            If Abs(v_temp(site)) = 1E+38 Then v_temp(site) = 0
            rtnYfup2(site) = StdCM_GetMax(Abs(h_temp(site)), Abs(v_temp(site))) / pCoef
        End If
    Next site

EndYfrm2:
    Call hDifPlane.GetSharedFlagPlane("FLG_YFRAME2_HV").RemoveFlagBit("FLG_YFRAME2_HV")
    Call hDifPlane.GetSharedFlagPlane("FLG_YFRAME2_H").RemoveFlagBit("FLG_YFRAME2_H")
    Call hDifPlane.GetSharedFlagPlane("FLG_YFRAME2_V").RemoveFlagBit("FLG_YFRAME2_V")
    Call hDifPlane.GetSharedFlagPlane("FLG_YFRAME2").RemoveFlagBit("FLG_YFRAME2")
    Call ReleasePlane(hDifPlane)
    Call ReleasePlane(vDifPlane)

End Sub

Public Sub std_CalcYglob( _
    ByRef localPlane As CImgPlane, ByVal pSlice As Double, ByVal pDif As Double, ByVal pCoef As Double, ByRef rtnYglob() As Double, ByRef rtnYgUpp() As Double)

    Dim site As Long

    Dim globalPlane As CImgPlane
    Call GetFreePlane(globalPlane, "pyglobal", idpDepthF32, False, "global plane")
    Call MultiMean(localPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, globalPlane, "YGLOBAL_ZONE2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, COMP_YLOCAL_TO_YGLOBAL, COMP_YLOCAL_TO_YGLOBAL)
 
    Dim HiLimit(nSite) As Double, LoLimit(nSite) As Double
    Dim yLevel() As Double
    TheResult.GetResult "Y_MEAN", yLevel
    '///// Y MURA SLICE LEVEL ////////////////////
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            HiLimit(site) = yLevel(site) * pSlice * pCoef
            LoLimit(site) = HiLimit(site) * (-1)
        End If
    Next site

    Dim hDifPlane As CImgPlane, vDifPlane As CImgPlane
    Call GetFreePlane(hDifPlane, globalPlane.planeGroup, idpDepthF32, , "hDifPlane")
    Call GetFreePlane(vDifPlane, globalPlane.planeGroup, idpDepthF32, , "vDifPlane")
    
    Dim globalWorkPlane0 As CImgPlane
    Call GetFreePlane(globalWorkPlane0, globalPlane.planeGroup, idpDepthF32, False, "global work plane0")
    Call Copy(globalPlane, "YGLOBAL_COL_SUB_TARGET", EEE_COLOR_FLAT, globalWorkPlane0, "YGLOBAL_COL_SUB_TARGET", EEE_COLOR_FLAT)
    Call Subtract(globalPlane, "YGLOBAL_COL_SUB_SOURCE", EEE_COLOR_FLAT, _
                  globalWorkPlane0, "YGLOBAL_COL_SUB_TARGET", EEE_COLOR_FLAT, _
                  hDifPlane, "YGLOBAL_COL_SUB_SOURCE", EEE_COLOR_FLAT)
    Call ReleasePlane(globalWorkPlane0)

    Call SubRows(globalPlane, "YGLOBAL_ZONE2D", EEE_COLOR_FLAT, vDifPlane, "YGLOBAL_ZONE2D", EEE_COLOR_FLAT, pDif)

    Dim mYglobal(nSite) As Double, mYglobalH(nSite) As Double, mYglobalV(nSite) As Double
    '========== YGLOBAL COUNT =============================
    Call Count(hDifPlane, "YGLOBAL_COL_JUDGE", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mYglobalH, "FLG_YGLOBAL_H")
    Call Count(vDifPlane, "YGLOBAL_ROW_JUDGE", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mYglobalV, "FLG_YGLOBAL_V")
    '[IDV Point] "hDifPlane"=H Diff Plane (Flag="FLG_YGLOBAL_H")
    '[IDV Point] "vDifPlane"=V Diff Plane (Flag="FLG_YGLOBAL_V")
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            mYglobal(site) = mYglobalH(site) + mYglobalV(site)
        End If
    Next site

    '========== SKIP CHECK ================================
    Dim IsSkip As Boolean
    IsSkip = True
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If mYglobal(site) > 0 Then IsSkip = False
        End If
    Next site
    If IsSkip = True Then
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                rtnYglob(site) = 0
                rtnYgUpp(site) = 0
            End If
        Next site
        GoTo EndYglob
    End If

    Dim FlgPlane As CImgPlane
    Call GetFreePlane(FlgPlane, globalPlane.planeGroup, idpDepthS16, True, "flgPlane")
    '========== MAKE YGLOBAL DEFFECT IMAGE ================
    Call SharedFlagOr(FlgPlane.planeGroup, "YGLOBAL_ZONE2D", "FLG_YGLOBAL_HV", "FLG_YGLOBAL_H", "FLG_YGLOBAL_V")
    Call FlagCopy(FlgPlane, "YGLOBAL_ZONE2D", "FLG_YGLOBAL_HV")
    Call FlgPlane.GetSharedFlagPlane("FLG_YGLOBAL_HV").RemoveFlagBit("FLG_YGLOBAL_HV")
    Call FlgPlane.GetSharedFlagPlane("FLG_YGLOBAL_H").RemoveFlagBit("FLG_YGLOBAL_H")
    Call FlgPlane.GetSharedFlagPlane("FLG_YGLOBAL_V").RemoveFlagBit("FLG_YGLOBAL_V")

    Dim cfuPlane As CImgPlane
    Call GetFreePlane(cfuPlane, globalPlane.planeGroup, idpDepthS16, False, "cfuPlane")
    '========== CONVOLUTION ===============================
    Call Convolution(FlgPlane, "YGLOBAL_ZONE2D", EEE_COLOR_FLAT, cfuPlane, "YGLOBAL_ZONE2D", EEE_COLOR_FLAT, Kernel_YGlobal)
    Call ReleasePlane(FlgPlane)

    '========== COUNT YGLOBAL ==============================
    Call Count(cfuPlane, "YGLOBAL_ZONE2D", EEE_COLOR_FLAT, idpCountAbove, yGlobCount, yGlobCount, idpLimitExclude, rtnYglob, "FLG_YGLOBAL")
    '[IDV Point] "cfuPlane"=YGLOB Final Judgement Plane
    Call ReleasePlane(cfuPlane)

    Dim h_temp(nSite) As Double, v_temp(nSite) As Double
    Call AbsMax(hDifPlane, "YGLOBAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, h_temp, "FLG_YGLOBAL")
    Call AbsMax(vDifPlane, "YGLOBAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, v_temp, "FLG_YGLOBAL")
    '[IDV Point] "hDifPlane"=YG_UPP Judgement Plane (H Diff)(InputFlag="FLG_YGLOBAL")
    '[IDV Point] "vDifPlane"=YG_UPP Judgement Plane (V Diff)(InputFlag="FLG_YGLOBAL")
    Call globalPlane.GetSharedFlagPlane("FLG_YGLOBAL").RemoveFlagBit("FLG_YGLOBAL")
    Call ReleasePlane(hDifPlane)
    Call ReleasePlane(vDifPlane)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If Abs(h_temp(site)) = 1E+38 Then h_temp(site) = 0
            If Abs(v_temp(site)) = 1E+38 Then v_temp(site) = 0
            rtnYgUpp(site) = StdCM_GetMax(Abs(h_temp(site)), Abs(v_temp(site))) / pCoef
        End If
    Next site

EndYglob:
    Call globalPlane.GetSharedFlagPlane("FLG_YGLOBAL_HV").RemoveFlagBit("FLG_YGLOBAL_HV")
    Call globalPlane.GetSharedFlagPlane("FLG_YGLOBAL_H").RemoveFlagBit("FLG_YGLOBAL_H")
    Call globalPlane.GetSharedFlagPlane("FLG_YGLOBAL_V").RemoveFlagBit("FLG_YGLOBAL_V")
    Call globalPlane.GetSharedFlagPlane("FLG_YGLOBAL").RemoveFlagBit("FLG_YGLOBAL")
    Call ReleasePlane(globalPlane)

End Sub


Public Sub pre_CalcClocal( _
ByRef fLowPassRed As CImgPlane, ByRef fLowPassGreen As CImgPlane, ByRef fLowPassBlue As CImgPlane, ByRef gainedRyPlane As CImgPlane, ByRef gainedByPlane As CImgPlane)

    '画像をさらに1/2圧縮。
    Dim rLocalPlane As CImgPlane
    Dim gLocalPlane As CImgPlane
    Dim bLocalPlane As CImgPlane
    Call GetFreePlane(rLocalPlane, "pylocal", idpDepthF32, False, "local(R)")
    Call GetFreePlane(gLocalPlane, "pylocal", idpDepthF32, False, "local(G)")
    Call GetFreePlane(bLocalPlane, "pylocal", idpDepthF32, False, "local(B)")
    Call MultiMean(fLowPassRed, "YLINE_ZONE2D", EEE_COLOR_FLAT, rLocalPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, COMP_YLINE_TO_YLOCAL, COMP_YLINE_TO_YLOCAL)
    Call MultiMean(fLowPassGreen, "YLINE_ZONE2D", EEE_COLOR_FLAT, gLocalPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, COMP_YLINE_TO_YLOCAL, COMP_YLINE_TO_YLOCAL)
    Call MultiMean(fLowPassBlue, "YLINE_ZONE2D", EEE_COLOR_FLAT, bLocalPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, COMP_YLINE_TO_YLOCAL, COMP_YLINE_TO_YLOCAL)

    '色差信号作成
    Dim ryTmpPlane As CImgPlane
    Dim byTmpPlane As CImgPlane
    Call GetFreePlane(ryTmpPlane, rLocalPlane.planeGroup, idpDepthF32, , "ryTmpPlane")
    Call GetFreePlane(byTmpPlane, rLocalPlane.planeGroup, idpDepthF32, , "byTmpPlane")
    Call StdCM_MakeRyBy(rLocalPlane, gLocalPlane, bLocalPlane, ryTmpPlane, byTmpPlane)
    Call ReleasePlane(rLocalPlane)
    Call ReleasePlane(gLocalPlane)
    Call ReleasePlane(bLocalPlane)

    Dim ryPlane As CImgPlane
    Dim byPlane As CImgPlane
    Call GetFreePlane(ryPlane, "pclocal", idpDepthF32, False, "ryPlane")
    Call GetFreePlane(byPlane, "pclocal", idpDepthF32, False, "byPlane")
    Call MultiMean(ryTmpPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, ryPlane, "CLOCAL_ZONE2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, COMP_YLOCAL_TO_CLOCAL, COMP_YLOCAL_TO_CLOCAL)
    Call MultiMean(byTmpPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, byPlane, "CLOCAL_ZONE2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, COMP_YLOCAL_TO_CLOCAL, COMP_YLOCAL_TO_CLOCAL)
    '[IDV Point] "ryPlane"= RY
    '[IDV Point] "byPlane"= BY
    Call ReleasePlane(ryTmpPlane)
    Call ReleasePlane(byTmpPlane)

    '========== Get DEGREE DATA ===========================
    '色相を求める。
    Dim degPlane As CImgPlane
    Call GetFreePlane(degPlane, ryPlane.planeGroup, idpDepthS16, , "degPlane")
    Call StdCM_DegcalNew(ryPlane, byPlane, degPlane, "CLOCAL_ZONE2D")
    '[IDV Point] "degPlane"= Hue (in degree)

    '========== CORRECT LEVEL =============================
    '色相を補正係数に変換。
    Dim hueGainPlane As CImgPlane
    Call GetFreePlane(hueGainPlane, degPlane.planeGroup, idpDepthS16, False, "hue gain")
    Call ExecuteLUT(degPlane, "CLOCAL_ZONE2D", EEE_COLOR_FLAT, hueGainPlane, "CLOCAL_ZONE2D", EEE_COLOR_FLAT, lookUpTable_lut15)
    
    Dim clocalWorkPlane0 As CImgPlane
    Dim clocalWorkPlane1 As CImgPlane
    Call GetFreePlane(clocalWorkPlane0, ryPlane.planeGroup, idpDepthF32, False, "clocal work plane0")
    Call GetFreePlane(clocalWorkPlane1, ryPlane.planeGroup, idpDepthF32, False, "clocal work plane1")
    Call Copy(hueGainPlane, "CLOCAL_ZONE2D", EEE_COLOR_FLAT, clocalWorkPlane0, "CLOCAL_ZONE2D", EEE_COLOR_FLAT)
    Call DivideConst(clocalWorkPlane0, "CLOCAL_ZONE2D", EEE_COLOR_FLAT, 256, clocalWorkPlane1, "CLOCAL_ZONE2D", EEE_COLOR_FLAT)
    Call ReleasePlane(degPlane)
    Call ReleasePlane(hueGainPlane)

'    Call GetFreePlane(gainedRyPlane, ryPlane.PlaneGroup, idpDepthF32, , "gained Ry plane")
'    Call GetFreePlane(gainedByPlane, ryPlane.PlaneGroup, idpDepthF32, , "gained By plane")
    '========== DEGREE * RY,BY DATA =======================
    Call Multiply(ryPlane, "CLOCAL_ZONE2D", EEE_COLOR_FLAT, _
                  clocalWorkPlane1, "CLOCAL_ZONE2D", EEE_COLOR_FLAT, _
                  gainedRyPlane, "CLOCAL_ZONE2D", EEE_COLOR_FLAT)
    Call Multiply(byPlane, "CLOCAL_ZONE2D", EEE_COLOR_FLAT, _
                  clocalWorkPlane1, "CLOCAL_ZONE2D", EEE_COLOR_FLAT, _
                  gainedByPlane, "CLOCAL_ZONE2D", EEE_COLOR_FLAT)
    '[IDV Point] "gained Ry plane"= RY with hue Correction
    '[IDV Point] "gained By plane"= BY with hue Correction

End Sub

Public Sub std_CalcClocal(ByRef ryHDifPlane As CImgPlane, ByRef ryVDifPlane As CImgPlane, ByRef rtnClocal() As Double)

    Dim site As Long

    Dim h_temp(nSite) As Double, v_temp(nSite) As Double, hv_temp(nSite) As Double
    Call max(ryHDifPlane, "CLOCAL_COL_JUDGE", EEE_COLOR_FLAT, h_temp)
    Call max(ryVDifPlane, "CLOCAL_ROW_JUDGE", EEE_COLOR_FLAT, v_temp)
    '[IDV Point] "ryHDifPlane"= (CLOCAL H-Diff Judgement Image)^2
    '[IDV Point] "ryVDifPlane"= (CLOCAL V-Diff Judgement Image)^2
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then

            '求めた最大値を"std_MakeRyBy"関数内で掛けていた"BitResolution"で割り戻し、昔の値と整合を取る"MagicNumber"
            'を乗じた上でムラ基準サイズ相当のレベルに戻します。
            hv_temp(site) = StdCM_GetMax(h_temp(site), v_temp(site))
            rtnClocal(site) = Sqr(hv_temp(site)) / cLocalCoef

        End If
    Next site

End Sub


Public Sub std_CalcCfrm2(ByRef ryHDifPlane As CImgPlane, ByRef ryVDifPlane As CImgPlane, ByRef rtnCfrm2() As Double)
    Dim site As Long
    
    '-------------------- FRAME2 --------------------------
    Dim h_temp(nSite) As Double, v_temp(nSite) As Double, hv_temp(nSite) As Double
    Call max(ryHDifPlane, "CLOCAL_ZONE2", EEE_COLOR_FLAT, h_temp)
    Call max(ryVDifPlane, "CLOCAL_ZONE2", EEE_COLOR_FLAT, v_temp)
    '[IDV Point] "hDif32Plane"= (CFRM2 H-Diff Judgement Image)^2
    '[IDV Point] "vDif32Plane"= (CFRM2 V-Diff Judgement Image)^2
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            '求めた最大値を"std_MakeRyBy"関数内で掛けていた"BitResolution"で割り戻し、昔の値と整合を取る"MagicNumber"
            'を乗じた上でムラ基準サイズ相当のレベルに戻します。
            hv_temp(site) = StdCM_GetMax(h_temp(site), v_temp(site))
            rtnCfrm2(site) = Sqr(hv_temp(site)) / cLocalCoef
        End If
    Next site

End Sub

Public Sub std_CalcCfrm2d(ByRef ryHDifPlane As CImgPlane, ByRef ryVDifPlane As CImgPlane, ByRef rtnCfrm2d() As Double)
    
    Dim site As Long
    'ZONE2をゼロクリア。
    Call WritePixel(ryHDifPlane, "CLOCAL_ZONE2", EEE_COLOR_FLAT, 0)
    Call WritePixel(ryVDifPlane, "CLOCAL_ZONE2", EEE_COLOR_FLAT, 0)

    '-------------------- FRAM2D --------------------------
    Dim h_temp(nSite) As Double, v_temp(nSite) As Double, hv_temp(nSite) As Double
    Call max(ryHDifPlane, "CLOCAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, h_temp)
    Call max(ryVDifPlane, "CLOCAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, v_temp)
    '[IDV Point] "hDif32Plane"= (CFRM2D H-Diff Judgement Image)^2
    '[IDV Point] "vDif32Plane"= (CFRM2D V-Diff Judgement Image)^2
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            '求めた最大値を"std_MakeRyBy"関数内で掛けていた"BitResolution"で割り戻し、昔の値と整合を取る"MagicNumber"
            'を乗じた上でムラ基準サイズ相当のレベルに戻します。
            hv_temp(site) = StdCM_GetMax(h_temp(site), v_temp(site))
            rtnCfrm2d(site) = Sqr(hv_temp(site)) / cLocalCoef
            
        End If
    Next site

End Sub

Public Sub std_CalcCglob01(ByRef cglo_ryHDifPlane As CImgPlane, ByRef cglo_ryVDifPlane As CImgPlane, ByRef rtnCglo1() As Double)
    
    Dim site As Long
    
    Dim h_temp(nSite) As Double, v_temp(nSite) As Double, hv_temp(nSite) As Double
    Call max(cglo_ryHDifPlane, "CGLOBAL_COL_JUDGE", EEE_COLOR_FLAT, h_temp)
    Call max(cglo_ryVDifPlane, "CGLOBAL_ROW_JUDGE", EEE_COLOR_FLAT, v_temp)
    '[IDV Point] "ryHDifPlane"= (CGLOB H-Diff Judgement Image)^2
    '[IDV Point] "ryVDifPlane"= (CGLOB V-Diff Judgement Image)^2
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            '求めた最大値を"std_MakeRyBy"関数内で掛けていた"BitResolution"で割り戻し、昔の値と整合を取る"MagicNumber"
            'を乗じた上でムラ基準サイズ相当のレベルに戻します。
            hv_temp(site) = StdCM_GetMax(h_temp(site), v_temp(site))
            rtnCglo1(site) = Sqr(hv_temp(site)) / cGlobCoef
        End If
    Next site
End Sub

Public Sub std_CalcCglob2d(ByRef cglo_ryHDifPlane As CImgPlane, ByRef cglo_ryVDifPlane As CImgPlane, ByRef rtnCglo2d() As Double)
    
    Dim site As Long
    
    Dim h_temp(nSite) As Double, v_temp(nSite) As Double, hv_temp(nSite) As Double

    Call WritePixel(cglo_ryHDifPlane, "CGLOBAL_COL_JUDGE", EEE_COLOR_FLAT, 0)
    Call WritePixel(cglo_ryVDifPlane, "CGLOBAL_ROW_JUDGE", EEE_COLOR_FLAT, 0)

    Call max(cglo_ryHDifPlane, "CGLOBAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, h_temp)
    Call max(cglo_ryVDifPlane, "CGLOBAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, v_temp)
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            '求めた最大値を"std_MakeRyBy"関数内で掛けていた"BitResolution"で割り戻し、昔の値と整合を取る"MagicNumber"
            'を乗じた上でムラ基準サイズ相当のレベルに戻します。
            hv_temp(site) = StdCM_GetMax(h_temp(site), v_temp(site))
            rtnCglo2d(site) = Sqr(hv_temp(site)) / cGlobCoef
            
        End If
    Next site

End Sub

Public Sub std_CalcCshad(ByRef ryGainedPlane As CImgPlane, ByRef byGainedPlane As CImgPlane, ByRef rtnCshad() As Double)

    Dim site As Long

    '>>>> Plane Info >>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    'ryGainedPlane     :R-Y(色相補正済み)
    'byGainedPlane     :B-Y(色相補正済み)
    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    'R-YとB-Yの最大・最小のアドレスにそれぞれフラグを立てる。
    Dim R_max(nSite) As Double, R_min(nSite) As Double, B_max(nSite) As Double, B_min(nSite) As Double
    Call MinMax(ryGainedPlane, "CGLOBAL_ZONE2D", EEE_COLOR_FLAT, R_min, R_max)
    Call MinMax(byGainedPlane, "CGLOBAL_ZONE2D", EEE_COLOR_FLAT, B_min, B_max)

    Call PutFlag(ryGainedPlane, "CGLOBAL_ZONE2D", EEE_COLOR_FLAT, idpCountBetween, R_min, R_min, idpLimitInclude, "RY_MIN")
    Call PutFlag(ryGainedPlane, "CGLOBAL_ZONE2D", EEE_COLOR_FLAT, idpCountBetween, R_max, R_max, idpLimitInclude, "RY_MAX")
    Call PutFlag(byGainedPlane, "CGLOBAL_ZONE2D", EEE_COLOR_FLAT, idpCountBetween, B_min, B_min, idpLimitInclude, "BY_MIN")
    Call PutFlag(byGainedPlane, "CGLOBAL_ZONE2D", EEE_COLOR_FLAT, idpCountBetween, B_max, B_max, idpLimitInclude, "BY_MAX")

    'フラグ位置に対応する補正後のR-Y,B-Y値を求め、色相・色差平面での最大距離を求める。
    'メモ：サイトシリアルループをまわすよりも、MinやMaxを用いて直接値を読むほうが、複数個取りの場合
    '高速である(シミュレーターではシリアルループ方式: 21msec程度に対して、Min/Max利用方式: 28msec
    '@Singleなので、2個取り以上の場合、Min/Max方式のほうが速くなる)。しかし、もし上のPutFlagで
    '2個以上のフラグが立ったとき(非常にまれだが、同じ最大値・最小値をとるピクセルが2個以上存在する場合)、
    '本来求めるべき距離が得られない恐れがあるため、シリアルループ方式にする。
    Dim tmpPixLog() As T_PIXINFO
    Dim r_RMin(nSite) As Double, r_RMax(nSite) As Double, b_RMin(nSite) As Double, b_RMax(nSite) As Double
    Dim r_BMin(nSite) As Double, r_BMax(nSite) As Double, b_BMin(nSite) As Double, b_BMax(nSite) As Double
    Dim R_vec(nSite) As Double, B_vec(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            'RY最小位置の(RY,BY)
            Call ReadPixel(ryGainedPlane, "CGLOBAL_ZONE2D", site, 1, "RY_MIN", tmpPixLog, idpAddrAbsolute)
            r_RMin(site) = tmpPixLog(0).Value
            Call ReadPixel(byGainedPlane, "CGLOBAL_ZONE2D", site, 1, "RY_MIN", tmpPixLog, idpAddrAbsolute)
            b_RMin(site) = tmpPixLog(0).Value

            'RY最大位置の(RY,BY)
            Call ReadPixel(ryGainedPlane, "CGLOBAL_ZONE2D", site, 1, "RY_MAX", tmpPixLog, idpAddrAbsolute)
            r_RMax(site) = tmpPixLog(0).Value
            Call ReadPixel(byGainedPlane, "CGLOBAL_ZONE2D", site, 1, "RY_MAX", tmpPixLog, idpAddrAbsolute)
            b_RMax(site) = tmpPixLog(0).Value

            'BY最小位置の(RY,BY)
            Call ReadPixel(ryGainedPlane, "CGLOBAL_ZONE2D", site, 1, "BY_MIN", tmpPixLog, idpAddrAbsolute)
            r_BMin(site) = tmpPixLog(0).Value
            Call ReadPixel(byGainedPlane, "CGLOBAL_ZONE2D", site, 1, "BY_MIN", tmpPixLog, idpAddrAbsolute)
            b_BMin(site) = tmpPixLog(0).Value

            'BY最大位置の(RY,BY)
            Call ReadPixel(ryGainedPlane, "CGLOBAL_ZONE2D", site, 1, "BY_MAX", tmpPixLog, idpAddrAbsolute)
            r_BMax(site) = tmpPixLog(0).Value
            Call ReadPixel(byGainedPlane, "CGLOBAL_ZONE2D", site, 1, "BY_MAX", tmpPixLog, idpAddrAbsolute)
            b_BMax(site) = tmpPixLog(0).Value


            '色差位置ベクトルの最大最小間距離を計算(ピタゴラスの定理から)。
            R_vec(site) = Sqr((r_RMax(site) - r_RMin(site)) ^ 2 + (b_RMax(site) - b_RMin(site)) ^ 2)
            B_vec(site) = Sqr((r_BMax(site) - r_BMin(site)) ^ 2 + (b_BMax(site) - b_BMin(site)) ^ 2)
            rtnCshad(site) = StdCM_GetMax(R_vec(site), B_vec(site))

        End If
    Next site

    Call ryGainedPlane.GetSharedFlagPlane("RY_MIN").RemoveFlagBit("RY_MIN")
    Call ryGainedPlane.GetSharedFlagPlane("RY_MAX").RemoveFlagBit("RY_MAX")
    Call ryGainedPlane.GetSharedFlagPlane("BY_MIN").RemoveFlagBit("BY_MIN")
    Call ryGainedPlane.GetSharedFlagPlane("BY_MAX").RemoveFlagBit("BY_MAX")


End Sub

Public Sub std_TCalcYline( _
     ByRef yPlane As CImgPlane, ByVal pSlice As Double, ByVal pDif As Double, ByVal pCoef As Double, ByRef rtnYline() As Double)

    Dim site As Long

    Dim HiLimit(nSite) As Double, LoLimit(nSite) As Double
    Dim yLevel() As Double
    TheResult.GetResult "Y_MEAN", yLevel
    '///// Y MURA SLICE LEVEL /////////////////////////////
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            HiLimit(site) = yLevel(site) * pSlice * pCoef
            LoLimit(site) = HiLimit(site) * (-1)
        End If
    Next site

    Dim hDifPlane As CImgPlane, vDifPlane As CImgPlane
    Call GetFreePlane(hDifPlane, yPlane.planeGroup, idpDepthF32, , "hDifPlane")
    Call GetFreePlane(vDifPlane, yPlane.planeGroup, idpDepthF32, , "vDifPlane")

    '========== H/V LINE DIFF.============================
    Dim workPlane1 As CImgPlane
    Call GetFreePlane(workPlane1, yPlane.planeGroup, idpDepthF32)
    Call Copy(yPlane, "YLINE_COL_SUB_TARGET_L", EEE_COLOR_FLAT, workPlane1, "YLINE_COL_SUB_TARGET_L", EEE_COLOR_FLAT)
    Call Subtract(yPlane, "YLINE_COL_SUB_SOURCE_L", EEE_COLOR_FLAT, _
                  workPlane1, "YLINE_COL_SUB_TARGET_L", EEE_COLOR_FLAT, _
                  hDifPlane, "YLINE_COL_SUB_SOURCE_L", EEE_COLOR_FLAT)
    Call ReleasePlane(workPlane1)

    Dim workPlane2 As CImgPlane
    Call GetFreePlane(workPlane2, yPlane.planeGroup, idpDepthF32)
    Call Copy(yPlane, "YLINE_COL_SUB_TARGET_R", EEE_COLOR_FLAT, workPlane2, "YLINE_COL_SUB_TARGET_R", EEE_COLOR_FLAT)
    Call Subtract(yPlane, "YLINE_COL_SUB_SOURCE_R", EEE_COLOR_FLAT, _
                  workPlane2, "YLINE_COL_SUB_TARGET_R", EEE_COLOR_FLAT, _
                  hDifPlane, "YLINE_COL_SUB_SOURCE_R", EEE_COLOR_FLAT)
    Call ReleasePlane(workPlane2)

    Call SubRows(yPlane, "YLINE_ROW_SUB", EEE_COLOR_FLAT, vDifPlane, "YLINE_ROW_SUB", EEE_COLOR_FLAT, pDif)

    Dim mYlineH(nSite) As Double, mYlineV(nSite) As Double, mYline(nSite) As Double
    '========== YLINE COUNT ===============================
    Call Count(hDifPlane, "YLINE_COL_JUDGE", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mYlineH, "FLG_YLINE_H")
    Call Count(vDifPlane, "YLINE_ROW_JUDGE", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mYlineV, "FLG_YLINE_V")
    '[IDV Point] "hDifPlane"=H Diff Plane (Flag="FLG_YLINE_H")
    '[IDV Point] "vDifPlane"=V Diff Plane (Flag="FLG_YLINE_V")
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            mYline(site) = mYlineH(site) + mYlineV(site)
        End If
    Next site
    
    Call ReleasePlane(hDifPlane)
    Call ReleasePlane(vDifPlane)

    '========== SKIP CHECK ================================
    Dim IsSkip As Boolean
    IsSkip = True
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If mYline(site) > 0 Then IsSkip = False
        End If
    Next site
    If IsSkip = True Then
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                rtnYline(site) = 0
            End If
        Next site
        GoTo EndYLine
    End If

    Dim FlgPlane As CImgPlane
    Call GetFreePlane(FlgPlane, yPlane.planeGroup, idpDepthS16, True, "flgPlane")
    '========== MAKE YLINE DEFFECT IMAGE ==================
    Call SharedFlagOr(FlgPlane.planeGroup, "YLINE_ZONE2D", "FLG_YLINE_HV", "FLG_YLINE_H", "FLG_YLINE_V")
    Call FlagCopy(FlgPlane, "YLINE_ZONE2D", "FLG_YLINE_HV")
    Call FlgPlane.GetSharedFlagPlane("FLG_YLINE_HV").RemoveFlagBit("FLG_YLINE_HV")
    Call FlgPlane.GetSharedFlagPlane("FLG_YLINE_H").RemoveFlagBit("FLG_YLINE_H")
    Call FlgPlane.GetSharedFlagPlane("FLG_YLINE_V").RemoveFlagBit("FLG_YLINE_V")

    Dim tmpPlane As CImgPlane
    Call GetFreePlane(tmpPlane, yPlane.planeGroup, idpDepthS16, True, "tmpPlane")
    '========== 5BIT MEDIAN ===============================
    Call Median(FlgPlane, "YLINE_FULL", EEE_COLOR_FLAT, tmpPlane, "YLINE_FULL", EEE_COLOR_FLAT, 5, 1)
    Call Median(tmpPlane, "YLINE_FULL", EEE_COLOR_FLAT, FlgPlane, "YLINE_FULL", EEE_COLOR_FLAT, 1, 5)
    Call ReleasePlane(tmpPlane)

    Dim compPlane1 As CImgPlane
    Call GetFreePlane(compPlane1, "pylinecomp", idpDepthS16, False, "compPlane")
    '========== MULTIMEAN =================================
    Call MultiMean(FlgPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, compPlane1, "YLINECOMP_FULL", EEE_COLOR_FLAT, idpMultiMeanFuncSum, yLineCompPix, yLineCompPix)
    Call ReleasePlane(FlgPlane)

    Dim rtnVal(nSite) As Double
    Dim compPlane2 As CImgPlane
    Call GetFreePlane(compPlane2, compPlane1.planeGroup, idpDepthS16, False, "comp convolution plane")
    Call CountForFlgBitImgPlane(compPlane1, "YLINECOMP_FULL", EEE_COLOR_FLAT, idpCountAbove, 1, 1, idpLimitExclude, rtnVal, compPlane2, 1)

    Call Convolution(compPlane2, "YLINECOMP_FULL", EEE_COLOR_FLAT, compPlane1, "YLINECOMP_FULL", EEE_COLOR_FLAT, Kernel_YLine)

    '========== COUNT YLINE ===============================
    Call Count(compPlane1, "YLINECOMP_FULL", EEE_COLOR_FLAT, idpCountAbove, yLineCount, yLineCount, idpLimitExclude, rtnYline)
    '[IDV Point] "compPlane"=YLINE Final Judgement Plane
    Call ReleasePlane(compPlane1)
    Call ReleasePlane(compPlane2)

EndYLine:
    Call yPlane.GetSharedFlagPlane("FLG_YLINE_HV").RemoveFlagBit("FLG_YLINE_HV")
    Call yPlane.GetSharedFlagPlane("FLG_YLINE_H").RemoveFlagBit("FLG_YLINE_H")
    Call yPlane.GetSharedFlagPlane("FLG_YLINE_V").RemoveFlagBit("FLG_YLINE_V")

End Sub

Public Sub std_TCalcYglob( _
    ByRef localPlane As CImgPlane, ByVal pSlice As Double, ByVal pDif As Double, ByVal pCoef As Double, ByRef rtnYglob() As Double, ByRef rtnYgUpp() As Double)

    Dim site As Long

    Dim globalPlane As CImgPlane
    Call GetFreePlane(globalPlane, "pyglobal", idpDepthF32, False, "global plane")
    Call MultiMean(localPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, globalPlane, "YGLOBAL_ZONE2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, COMP_YLOCAL_TO_YGLOBAL, COMP_YLOCAL_TO_YGLOBAL)
 
    Dim HiLimit(nSite) As Double, LoLimit(nSite) As Double
    Dim yLevel() As Double
    TheResult.GetResult "Y_MEAN", yLevel
    '///// Y MURA SLICE LEVEL ////////////////////
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            HiLimit(site) = yLevel(site) * pSlice * pCoef
            LoLimit(site) = HiLimit(site) * (-1)
        End If
    Next site

    Dim hDifPlane As CImgPlane, vDifPlane As CImgPlane
    Call GetFreePlane(hDifPlane, globalPlane.planeGroup, idpDepthF32, , "hDifPlane")
    Call GetFreePlane(vDifPlane, globalPlane.planeGroup, idpDepthF32, , "vDifPlane")
    
    Dim globalWorkPlane0 As CImgPlane
    Call GetFreePlane(globalWorkPlane0, globalPlane.planeGroup, idpDepthF32, False, "global work plane0")
    Call Copy(globalPlane, "YGLOBAL_COL_SUB_TARGET_L", EEE_COLOR_FLAT, globalWorkPlane0, "YGLOBAL_COL_SUB_TARGET_L", EEE_COLOR_FLAT)
    Call Subtract(globalPlane, "YGLOBAL_COL_SUB_SOURCE_L", EEE_COLOR_FLAT, _
                  globalWorkPlane0, "YGLOBAL_COL_SUB_TARGET_L", EEE_COLOR_FLAT, _
                  hDifPlane, "YGLOBAL_COL_SUB_SOURCE_L", EEE_COLOR_FLAT)
    Call ReleasePlane(globalWorkPlane0)

    Dim globalWorkPlane1 As CImgPlane
    Call GetFreePlane(globalWorkPlane1, globalPlane.planeGroup, idpDepthF32, False, "global work plane1")
    Call Copy(globalPlane, "YGLOBAL_COL_SUB_TARGET_R", EEE_COLOR_FLAT, globalWorkPlane1, "YGLOBAL_COL_SUB_TARGET_R", EEE_COLOR_FLAT)
    Call Subtract(globalPlane, "YGLOBAL_COL_SUB_SOURCE_R", EEE_COLOR_FLAT, _
                  globalWorkPlane1, "YGLOBAL_COL_SUB_TARGET_R", EEE_COLOR_FLAT, _
                  hDifPlane, "YGLOBAL_COL_SUB_SOURCE_R", EEE_COLOR_FLAT)
    Call ReleasePlane(globalWorkPlane1)

    Call SubRows(globalPlane, "YGLOBAL_ZONE2D", EEE_COLOR_FLAT, vDifPlane, "YGLOBAL_ZONE2D", EEE_COLOR_FLAT, pDif)

    Dim mYglobal(nSite) As Double, mYglobalH(nSite) As Double, mYglobalV(nSite) As Double
    '========== YGLOBAL COUNT =============================
    Call Count(hDifPlane, "YGLOBAL_COL_JUDGE", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mYglobalH, "FLG_YGLOBAL_H")
    Call Count(vDifPlane, "YGLOBAL_ROW_JUDGE", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mYglobalV, "FLG_YGLOBAL_V")
    '[IDV Point] "hDifPlane"=H Diff Plane (Flag="FLG_YGLOBAL_H")
    '[IDV Point] "vDifPlane"=V Diff Plane (Flag="FLG_YGLOBAL_V")
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            mYglobal(site) = mYglobalH(site) + mYglobalV(site)
        End If
    Next site

    '========== SKIP CHECK ================================
    Dim IsSkip As Boolean
    IsSkip = True
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If mYglobal(site) > 0 Then IsSkip = False
        End If
    Next site
    If IsSkip = True Then
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                rtnYglob(site) = 0
                rtnYgUpp(site) = 0
            End If
        Next site
        GoTo EndYglob
    End If

    Dim FlgPlane As CImgPlane
    Call GetFreePlane(FlgPlane, globalPlane.planeGroup, idpDepthS16, True, "flgPlane")
    '========== MAKE YGLOBAL DEFFECT IMAGE ================
    Call SharedFlagOr(FlgPlane.planeGroup, "YGLOBAL_ZONE2D", "FLG_YGLOBAL_HV", "FLG_YGLOBAL_H", "FLG_YGLOBAL_V")
    Call FlagCopy(FlgPlane, "YGLOBAL_ZONE2D", "FLG_YGLOBAL_HV")
    Call FlgPlane.GetSharedFlagPlane("FLG_YGLOBAL_HV").RemoveFlagBit("FLG_YGLOBAL_HV")
    Call FlgPlane.GetSharedFlagPlane("FLG_YGLOBAL_H").RemoveFlagBit("FLG_YGLOBAL_H")
    Call FlgPlane.GetSharedFlagPlane("FLG_YGLOBAL_V").RemoveFlagBit("FLG_YGLOBAL_V")

    Dim cfuPlane As CImgPlane
    Call GetFreePlane(cfuPlane, globalPlane.planeGroup, idpDepthS16, False, "cfuPlane")
    '========== CONVOLUTION ===============================
    Call Convolution(FlgPlane, "YGLOBAL_ZONE2D", EEE_COLOR_FLAT, cfuPlane, "YGLOBAL_ZONE2D", EEE_COLOR_FLAT, Kernel_YGlobal)
    Call ReleasePlane(FlgPlane)

    '========== COUNT YGLOBAL ==============================
    Call Count(cfuPlane, "YGLOBAL_ZONE2D", EEE_COLOR_FLAT, idpCountAbove, yGlobCount, yGlobCount, idpLimitExclude, rtnYglob, "FLG_YGLOBAL")
    '[IDV Point] "cfuPlane"=YGLOB Final Judgement Plane
    Call ReleasePlane(cfuPlane)

    Dim h_temp(nSite) As Double, v_temp(nSite) As Double
    Call AbsMax(hDifPlane, "YGLOBAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, h_temp, "FLG_YGLOBAL")
    Call AbsMax(vDifPlane, "YGLOBAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, v_temp, "FLG_YGLOBAL")
    '[IDV Point] "hDifPlane"=YG_UPP Judgement Plane (H Diff)(InputFlag="FLG_YGLOBAL")
    '[IDV Point] "vDifPlane"=YG_UPP Judgement Plane (V Diff)(InputFlag="FLG_YGLOBAL")
    Call globalPlane.GetSharedFlagPlane("FLG_YGLOBAL").RemoveFlagBit("FLG_YGLOBAL")
    Call ReleasePlane(hDifPlane)
    Call ReleasePlane(vDifPlane)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If Abs(h_temp(site)) = 1E+38 Then h_temp(site) = 0
            If Abs(v_temp(site)) = 1E+38 Then v_temp(site) = 0
            rtnYgUpp(site) = StdCM_GetMax(Abs(h_temp(site)), Abs(v_temp(site))) / pCoef
        End If
    Next site

EndYglob:
    Call globalPlane.GetSharedFlagPlane("FLG_YGLOBAL_HV").RemoveFlagBit("FLG_YGLOBAL_HV")
    Call globalPlane.GetSharedFlagPlane("FLG_YGLOBAL_H").RemoveFlagBit("FLG_YGLOBAL_H")
    Call globalPlane.GetSharedFlagPlane("FLG_YGLOBAL_V").RemoveFlagBit("FLG_YGLOBAL_V")
    Call globalPlane.GetSharedFlagPlane("FLG_YGLOBAL").RemoveFlagBit("FLG_YGLOBAL")
    Call ReleasePlane(globalPlane)

End Sub

Private Sub std_TCalcYglob2( _
    ByRef yPlane As CImgPlane, ByVal pSlice As Double, ByRef pHdif As Double, ByRef pVdif As Double, ByRef rtnYglob2() As Double, ByRef rtnYgUp2() As Double)

    Dim site As Long
    
    Dim HiLimit(nSite) As Double, LoLimit(nSite) As Double
    Dim yLevel() As Double
    TheResult.GetResult "Y_MEAN", yLevel
    '///// Y MURA SLICE LEVEL ////////////////////
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            HiLimit(site) = (yLevel(site) * pSlice) ^ 2
            LoLimit(site) = HiLimit(site) * (-1)
        End If
    Next site

    Dim hDifPlane As CImgPlane, hDifPlane1 As CImgPlane, hDifPlane2 As CImgPlane, vDifPlane As CImgPlane
    Call GetFreePlane(hDifPlane, yPlane.planeGroup, idpDepthF32, , "hDifPlane")
    Call GetFreePlane(hDifPlane1, yPlane.planeGroup, idpDepthF32, , "hDifPlane1")
    Call GetFreePlane(hDifPlane2, yPlane.planeGroup, idpDepthF32, , "hDifPlane2")
    Call GetFreePlane(vDifPlane, yPlane.planeGroup, idpDepthF32, , "vDifPlane")
    Dim tmp1Plane As CImgPlane, tmp2Plane As CImgPlane, tmp3Plane As CImgPlane
    Call GetFreePlane(tmp1Plane, yPlane.planeGroup, idpDepthF32, , "tmp1Plane")
    Call GetFreePlane(tmp2Plane, yPlane.planeGroup, idpDepthF32, , "tmp2Plane")
    Call GetFreePlane(tmp3Plane, yPlane.planeGroup, idpDepthF32, , "tmp3Plane")

    Call Copy(yPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, tmp1Plane, "YLINE_ZONE2D", EEE_COLOR_FLAT)

    '/*---------------------- H DIF_LEFT ---------------------*/
    Call Subtract(yPlane, "YLINE_ZONE2D_YG2_H_L", EEE_COLOR_FLAT, tmp1Plane, "YLINE_ZONE2D_YG2_LSFT_L", EEE_COLOR_FLAT, tmp2Plane, "YLINE_ZONE2D_YG2_H_L", EEE_COLOR_FLAT)
    Call Subtract(yPlane, "YLINE_ZONE2D_YG2_H_L", EEE_COLOR_FLAT, tmp1Plane, "YLINE_ZONE2D_YG2_RSFT_L", EEE_COLOR_FLAT, tmp3Plane, "YLINE_ZONE2D_YG2_H_L", EEE_COLOR_FLAT)
    Call Add(tmp2Plane, "YLINE_ZONE2D_YG2_H_L", EEE_COLOR_FLAT, tmp3Plane, "YLINE_ZONE2D_YG2_H_L", EEE_COLOR_FLAT, hDifPlane, "YLINE_ZONE2D_YG2_H_L", EEE_COLOR_FLAT)
    Call DivideConst(hDifPlane, "YLINE_ZONE2D_YG2_H_L", EEE_COLOR_FLAT, 2, hDifPlane, "YLINE_ZONE2D_YG2_H_L", EEE_COLOR_FLAT)

    '/*---------------------- H DIF_RIGHT ---------------------*/
    Call Subtract(yPlane, "YLINE_ZONE2D_YG2_H_R", EEE_COLOR_FLAT, tmp1Plane, "YLINE_ZONE2D_YG2_LSFT_R", EEE_COLOR_FLAT, tmp2Plane, "YLINE_ZONE2D_YG2_H_R", EEE_COLOR_FLAT)
    Call Subtract(yPlane, "YLINE_ZONE2D_YG2_H_R", EEE_COLOR_FLAT, tmp1Plane, "YLINE_ZONE2D_YG2_RSFT_R", EEE_COLOR_FLAT, tmp3Plane, "YLINE_ZONE2D_YG2_H_R", EEE_COLOR_FLAT)
    Call Add(tmp2Plane, "YLINE_ZONE2D_YG2_H_R", EEE_COLOR_FLAT, tmp3Plane, "YLINE_ZONE2D_YG2_H_R", EEE_COLOR_FLAT, hDifPlane, "YLINE_ZONE2D_YG2_H_R", EEE_COLOR_FLAT)
    Call DivideConst(hDifPlane, "YLINE_ZONE2D_YG2_H_R", EEE_COLOR_FLAT, 2, hDifPlane, "YLINE_ZONE2D_YG2_H_R", EEE_COLOR_FLAT)

    '/*---------------------- H DIF LEFT + RIGHT---------------------*/
    Call Copy(hDifPlane1, "YLINE_ZONE2D_YG2_H_L", idpColorAll, hDifPlane, "YLINE_ZONE2D_YG2_H_L", idpColorAll)
    Call Copy(hDifPlane2, "YLINE_ZONE2D_YG2_H_R", idpColorAll, hDifPlane, "YLINE_ZONE2D_YG2_H_R", idpColorAll)
    Call ReleasePlane(hDifPlane1)
    Call ReleasePlane(hDifPlane2)

    '/*---------------------- V DIF ---------------------*/
    Call Subtract(yPlane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT, tmp1Plane, "YLINE_ZONE2D_YG2_TSFT", EEE_COLOR_FLAT, tmp2Plane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT)
    Call Subtract(yPlane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT, tmp1Plane, "YLINE_ZONE2D_YG2_BSFT", EEE_COLOR_FLAT, tmp3Plane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT)
    Call Add(tmp2Plane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT, tmp3Plane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT, vDifPlane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT)
    Call DivideConst(vDifPlane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT, 2, vDifPlane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT)

    Call ReleasePlane(tmp1Plane)
    Call ReleasePlane(tmp2Plane)
    Call ReleasePlane(tmp3Plane)

    Dim hvDifPlane As CImgPlane
    Call GetFreePlane(hvDifPlane, yPlane.planeGroup, idpDepthF32, True, "hvdifPlane")
    'X^2
    Call Multiply(hDifPlane, "YLINE_ZONE2D_YG2_H", EEE_COLOR_FLAT, hDifPlane, "YLINE_ZONE2D_YG2_H", EEE_COLOR_FLAT, hDifPlane, "YLINE_ZONE2D_YG2_H", EEE_COLOR_FLAT)
    'Y^2
    Call Multiply(vDifPlane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT, vDifPlane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT, vDifPlane, "YLINE_ZONE2D_YG2_V", EEE_COLOR_FLAT)
    'X^2+Y^2
    Call Add(hDifPlane, "YLINE_ZONE2D_YG2_HV", EEE_COLOR_FLAT, vDifPlane, "YLINE_ZONE2D_YG2_HV", EEE_COLOR_FLAT, hvDifPlane, "YLINE_ZONE2D_YG2_HV", EEE_COLOR_FLAT)

    Call PutFlag(hvDifPlane, "YLINE_ZONE2D_YG2_HV", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, "FLG_YGLOB2")

    Call ReleasePlane(hDifPlane)
    Call ReleasePlane(vDifPlane)

    Dim FlgPlane As CImgPlane, convPlane As CImgPlane
    Call GetFreePlane(FlgPlane, yPlane.planeGroup, idpDepthS16, True, "flgPlane")
    Call GetFreePlane(convPlane, yPlane.planeGroup, idpDepthS16, True, "convPlane")
    '========== CONTINUATION POINT DEFECT =================
    Call FlagCopy(FlgPlane, "YLINE_ZONE2D_YG2_HV", "FLG_YGLOB2")
    Call Convolution(FlgPlane, "YLINE_ZONE2D_YG2_HV", EEE_COLOR_FLAT, convPlane, "YLINE_ZONE2D_YG2_HV", EEE_COLOR_FLAT, "Kernel_YGlobal2")
    Call Count(convPlane, "YLINE_ZONE2D_YG2_HV", EEE_COLOR_FLAT, idpCountAbove, yGlob2Count, yGlob2Count, idpLimitExclude, rtnYglob2, "FLG_YGUP2")

    Dim tmpYgUp2(nSite) As Double
    Dim tmpMin(nSite) As Double, tmpMax(nSite) As Double
    Call MinMax(hvDifPlane, "YLINE_ZONE2D_YG2_HV", EEE_COLOR_FLAT, tmpMin, tmpMax, "FLG_YGUP2")
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            If rtnYglob2(site) > 0 And yLevel(site) > 0 Then
                If Abs(tmpMax(site)) > Abs(tmpMin(site)) Then
                    tmpYgUp2(site) = tmpMax(site)
                Else
                    tmpYgUp2(site) = tmpMin(site)
                End If
                rtnYgUp2(site) = Div(Sqr(Abs(tmpYgUp2(site))), yLevel(site), 999)
            Else
                rtnYgUp2(site) = 0
            End If
        End If
    Next site

    Call yPlane.GetSharedFlagPlane("FLG_YGLOB2").RemoveFlagBit("FLG_YGLOB2")
    Call yPlane.GetSharedFlagPlane("FLG_YGUP2").RemoveFlagBit("FLG_YGUP2")
End Sub

Public Function StdCM_View( _
    ByRef inPlane As CImgPlane, _
    Optional ByVal verboseOn As Boolean = False, _
    Optional ByRef multiplier As Double = 0, _
    Optional ByVal targetZone As String = "", _
    Optional ByVal commentStr As String = "", _
    Optional ByVal targetSite As Long = 0) As Boolean
'32Bit浮動小数画像は、IDVでMSBの変更ができず確認しにくいとの意見に対応するため、
'32Bit浮動小数画像を、いったん16Bit整数プレーンにコピーして、IDVで確認できるよう
'にした。関数内でブレークがかかるのでそのタイミングでIDVで見てほしい。
'■引数
'   inPlane     入力の32Bit浮動小数画像プレーン。
'   verboseOn   イミディエイトウインドウに画像に関するメッセージを出力するかどうかのフラグ。
'   multiplier  入力の画像を何倍かしてから16Bit整数プレーンにコピーするかを指定する。
'               指定がない場合はよきに計らう。
'   targetZone  よきに計らう際に用いるゾーン情報。指定がなければ、現在セットされているゾーンになる。
'   commentStr  イミディエイトに表示するコメント文
'   targetSite  表示したいサイトを指定する。

    Dim retAvg(nSite) As Integer
    Dim retMin(nSite) As Integer
    Dim retMax(nSite) As Integer
    Dim viewPlane As CImgPlane
    Set viewPlane = stdCM_PrepareIntPlane(inPlane, retMin, retAvg, retMax, multiplier, targetZone)
    
    '何らかのエラーにより、プレーンが戻ってこなかった場合は終了します。
    If viewPlane Is Nothing Then Exit Function
    
    If verboseOn Then
        'コメント文の作成
        Dim myComment As String
        If commentStr = "" Then
            myComment = "Plane Name = "
        Else
            myComment = commentStr & " = "
        End If
        
        Dim site As Long
        Dim strAvg As String
        Dim strMin As String
        Dim strMax As String
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                strAvg = strAvg & retAvg(site) & "(" & site & ") "
                strMin = strMin & retMin(site) & "(" & site & ") "
                strMax = strMax & retMax(site) & "(" & site & ") "
            End If
        Next site
        
        'ユーザーが画像を見る際に参考にする情報を、イミディエイトに出力します。
        Debug.Print "--------------------------------"
        Debug.Print myComment & viewPlane.Name
        Debug.Print "Min(Site) = " & strMin
        Debug.Print "Avg(Site) = " & strAvg
        Debug.Print "Max(Site) = " & strMax
        Debug.Print "--------------------------------"
    End If
    
    'ここで強制ブレークしますので、IDVで画像を見てください。
    Dim mySite As Long
    If (targetSite >= 0) And (targetSite <= nSite) Then
        mySite = targetSite
    End If
    Dim myLog As Long
    myLog = stdCM_returnMSB(CDbl(retMin(mySite)), CDbl(retMax(mySite)))
    
    'Please use IDV to view the image!
    Call StdCM_PlaneOnIDV(viewPlane, , , myLog, mySite)
'    Stop
    
    Call ReleasePlane(viewPlane)
    
    StdCM_View = True
    
End Function

Public Function StdCM_WBImage( _
    ByRef inRedPlane As CImgPlane, _
    ByRef inGreenPlane As CImgPlane, _
    ByRef inBluePlane As CImgPlane, _
    Optional ByVal targetSite As Long = 0) As Boolean
    'プロダクションモードでは実行しません(とまってしまうから)
    If Not TheExec.RunMode = runModeProduction Then '関数の中にいれる。
        Dim wbRedPlane As CImgPlane
        Dim wbGreenPlane As CImgPlane
        Dim wbBluePlane As CImgPlane
        Dim tmpMaxR(nSite) As Integer
        Dim tmpMinR(nSite) As Integer
        Dim tmpAvgR(nSite) As Integer
        Dim tmpMaxG(nSite) As Integer
        Dim tmpMinG(nSite) As Integer
        Dim tmpAvgG(nSite) As Integer
        Dim tmpMaxB(nSite) As Integer
        Dim tmpMinB(nSite) As Integer
        Dim tmpAvgB(nSite) As Integer
        
        Dim rPlane As CImgPlane
        Dim gPlane As CImgPlane
        Dim bPlane As CImgPlane
        
        'Red image
        Call GetFreePlane(rPlane, "rbayer", idpDepthS16, True, "Red WB")
        Set wbRedPlane = stdCM_PrepareIntPlane(inRedPlane, tmpMinR, tmpAvgR, tmpMaxR, 1000 / 0.3, "YLINE_ZONE2D")
        If wbRedPlane Is Nothing Then Exit Function
        Call Copy(wbRedPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, rPlane, "RBAYER_FULL", EEE_COLOR_FLAT)
        Call ReleasePlane(wbRedPlane)
        'Green image
        Call GetFreePlane(gPlane, "gbayer", idpDepthS16, True, "Green WB")
        Set wbGreenPlane = stdCM_PrepareIntPlane(inGreenPlane, tmpMaxG, tmpAvgG, tmpMaxG, 1000 / 0.59, "YLINE_ZONE2D")
        If wbGreenPlane Is Nothing Then Exit Function
        Call Copy(wbGreenPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, gPlane, "GBAYER_FULL", EEE_COLOR_FLAT)
        Call ReleasePlane(wbGreenPlane)
        'Blue image
        Call GetFreePlane(bPlane, "bbayer", idpDepthS16, True, "Blue WB")
        Set wbBluePlane = stdCM_PrepareIntPlane(inBluePlane, tmpMinB, tmpAvgB, tmpMaxB, 1000 / 0.11, "YLINE_ZONE2D")
        If wbBluePlane Is Nothing Then Exit Function
        Call Copy(wbBluePlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, bPlane, "BBAYER_FULL", EEE_COLOR_FLAT)
        Call ReleasePlane(wbBluePlane)
        
'        Debug.Print "-----------------------------------------"
'        Debug.Print "Mean value of each image approx 1000"
'        Debug.Print "Red   = " & rPlane.Name
'        Debug.Print "Green = " & gPlane.Name
'        Debug.Print "Blue  = " & bPlane.Name
'        Debug.Print "-----------------------------------------"
        
        Dim mySite As Long
        If (targetSite >= 0) And (targetSite <= nSite) Then
            mySite = targetSite
        End If
        Dim myLogR As Long
        Dim myLogG As Long
        Dim myLogB As Long
        Dim myMSB As Long
        myLogR = stdCM_returnMSB(CDbl(tmpMinR(mySite)), CDbl(tmpMaxR(mySite)))
        myLogG = stdCM_returnMSB(CDbl(tmpMinG(mySite)), CDbl(tmpMaxG(mySite)))
        myLogB = stdCM_returnMSB(CDbl(tmpMinB(mySite)), CDbl(tmpMaxB(mySite)))
        If myLogR > myLogG Then
            myMSB = myLogR
        Else
            myMSB = myLogG
        End If
        If myMSB < myLogB Then
            myMSB = myLogB
        End If
        Call StdCM_PlaneOnIDV(gPlane, rPlane, bPlane, myMSB, mySite)
        
        'Here View White Balance Image
        Stop
        
        Call ReleasePlane(rPlane)
        Call ReleasePlane(gPlane)
        Call ReleasePlane(bPlane)
        
    End If
    StdCM_WBImage = True
End Function

Private Function stdCM_PrepareIntPlane( _
    ByRef inPlane As CImgPlane, _
    ByRef retMin() As Integer, _
    ByRef retAvg() As Integer, _
    ByRef retMax() As Integer, _
    Optional ByRef multiplier As Double = 0, _
    Optional ByVal targetZone As String = "") As CImgPlane
    
    '指定がなければ整数プレーンにコピーしたとき、ゾーン平均値が1000になるように浮動小数プレーンに定数積算をします。
    Const DEFAULT_TARGET_MEAN As Double = 1000
    Const DEFAULT_MAX_VALUE As Double = 10000
    
    'もし32Bit浮動小数プレーンでなければ終了します。
    If inPlane.BitDepth <> idpDepthF32 Then
        Call MsgBox("Bit depth is not F32. Abort.")
        Exit Function
    End If
    
    '対象ゾーンを決めます。指定がなければ、入力プレーンの現在のプレーン設定に従います。
    Dim viewZone As String
    If targetZone = "" Then
        viewZone = inPlane.CurrentPMD.Name
    Else
        viewZone = targetZone
    End If
    
    '入力画像の最小値・最大値・平均値を算出します。
    Dim myAvg(nSite) As Double
    Dim myMin(nSite) As Double
    Dim myMax(nSite) As Double
    Call Average(inPlane, viewZone, EEE_COLOR_FLAT, myAvg)
    Call MinMax(inPlane, viewZone, EEE_COLOR_FLAT, myMin, myMax)
    
    '浮動小数画像に対する定数乗数を算出します(サイト毎)。指定がなければ、整数プレーンの平均値が1000になるように
    '決めます。
    Dim myMultiplier(nSite) As Double
    Dim site As Long
    If multiplier = 0 Then
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                myMultiplier(site) = stdCM_div(DEFAULT_TARGET_MEAN, myAvg(site), 10000)
            End If
        Next site
    Else
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                myMultiplier(site) = multiplier
            End If
        Next site
    End If
    
    '最後にユーザーに提示する、整数プレーンの統計値を算出します。
    Dim trtAvg(nSite) As Double
    Dim trtMin(nSite) As Double
    Dim trtMax(nSite) As Double
    Dim tmpMax As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            trtAvg(site) = myAvg(site) * myMultiplier(site)
            trtMin(site) = myMin(site) * myMultiplier(site)
            trtMax(site) = myMax(site) * myMultiplier(site)
            
            '16Bitプレーンの深度を超えてしまわないかどうかのチェック。
            If Abs(trtMin(site)) > Abs(trtMax(site)) Then
                tmpMax = Abs(trtMin(site))
            Else
                tmpMax = Abs(trtMax(site))
            End If
            
            If tmpMax > DEFAULT_MAX_VALUE Then
                myMultiplier(site) = myMultiplier(site) / (tmpMax / DEFAULT_MAX_VALUE)
                trtAvg(site) = myAvg(site) * myMultiplier(site)
                trtMin(site) = myMin(site) * myMultiplier(site)
                trtMax(site) = myMax(site) * myMultiplier(site)
            End If
            
            retAvg(site) = Int(trtAvg(site))
            retMin(site) = Int(trtMin(site))
            retMax(site) = Int(trtMax(site))
        End If
    Next site
    
    '定数乗算を行います。
    Dim tmpPlane As CImgPlane
    Call GetFreePlane(tmpPlane, inPlane.planeGroup, idpDepthF32, True, "temporary")
    Call MultiplyConst(inPlane, viewZone, EEE_COLOR_FLAT, myMultiplier, tmpPlane, viewZone, EEE_COLOR_FLAT)
    
    '整数プレーンにコピーします。
    Dim viewPlane As CImgPlane
    Call GetFreePlane(viewPlane, inPlane.planeGroup, idpDepthS16, True, "View Plane")
    Call Copy(tmpPlane, viewZone, EEE_COLOR_FLAT, viewPlane, viewZone, EEE_COLOR_FLAT)
    
    Call ReleasePlane(tmpPlane)
    
    Set stdCM_PrepareIntPlane = viewPlane
    
End Function

Private Function StdCM_PlaneOnIDV( _
    ByRef inPlane As CImgPlane, _
    Optional ByRef redPlane As CImgPlane, _
    Optional ByRef bluePlane As CImgPlane, _
    Optional ByRef inMsbBit As Long = 0, _
    Optional ByRef targetSite As Long = 0)
    
    With theidv
        'フォーム開く
        .OpenForm
        'プレーン設定(Green)
        .PlaneNameGreen = inPlane.Name
        If inPlane.BitDepth = idpDepthS16 Then
            If inMsbBit > 0 Then
                If inMsbBit <= 16 Then
                    .MsbGreen = inMsbBit
                Else
                    .MsbGreen = 16
                End If
            End If
        ElseIf inPlane.BitDepth = idpDepthS32 Then
            If inMsbBit > 0 Then
                If inMsbBit <= 32 Then
                    .MsbGreen = inMsbBit
                Else
                    .MsbGreen = 32
                End If
            End If
        End If
        'プレーン設定(Red)
        If Not redPlane Is Nothing Then
            .PlaneNameRed = redPlane.Name
            If redPlane.BitDepth = idpDepthS16 Then
                If inMsbBit > 0 Then
                    If inMsbBit <= 16 Then
                        .MsbRed = inMsbBit
                    Else
                        .MsbRed = 16
                    End If
                End If
            ElseIf redPlane.BitDepth = idpDepthS32 Then
                If inMsbBit > 0 Then
                    If inMsbBit <= 32 Then
                        .MsbRed = inMsbBit
                    Else
                        .MsbRed = 32
                    End If
                End If
            End If
        End If
        'プレーン設定(Blue)
        If Not bluePlane Is Nothing Then
            .PlaneNameBlue = bluePlane.Name
            If bluePlane.BitDepth = idpDepthS16 Then
                If inMsbBit > 0 Then
                    If inMsbBit <= 16 Then
                        .MsbBlue = inMsbBit
                    Else
                        .MsbBlue = 16
                    End If
                End If
            ElseIf bluePlane.BitDepth = idpDepthS32 Then
                If inMsbBit > 0 Then
                    If inMsbBit <= 32 Then
                        .MsbBlue = inMsbBit
                    Else
                        .MsbBlue = 32
                    End If
                End If
            End If
        End If
        
        'PMD設定
        .PMDName = inPlane.CurrentPmdName
        'ハイライト色変更
        .HilightColor = "Red"
        'サイト設定
        If (targetSite >= 0) And (targetSite <= nSite) Then
            .site = targetSite
        End If
        
        '変更有効化
        .Refresh
    End With
End Function

Private Function stdCM_returnMSB( _
    ByRef inMin As Double, ByRef inMax As Double) As Long
    Dim tmpMax As Double
    If Abs(inMin) > Abs(inMax) Then
        tmpMax = Abs(inMin)
    Else
        tmpMax = Abs(inMax)
    End If
    Dim tmpLog As Double
    If tmpMax > 0 Then
        tmpLog = Log(tmpMax) / Log(2)
        stdCM_returnMSB = Int(tmpLog)
        If tmpLog > stdCM_returnMSB Then
            stdCM_returnMSB = stdCM_returnMSB + 1
        End If
    End If
End Function

Public Sub CM_RGBseparate(InputPlane As CImgPlane, InputZone As Variant, ByRef result() As Double)

    Dim site As Long

    '========== BAYER SEPARATION ==================================
    Dim bayerRedPlane As CImgPlane
    Dim bayerGreenPlane As CImgPlane
    Dim bayerBluePlane As CImgPlane
    Call GetFreePlane(bayerRedPlane, "rbayer", idpDepthS16, False, "redPlane for bayer")
    Call GetFreePlane(bayerGreenPlane, "gbayer", idpDepthS16, False, "greenPlane for bayer")
    Call GetFreePlane(bayerBluePlane, "bbayer", idpDepthS16, False, "bluePlane for bayer")
    
    Dim tmpBayerPlane As CImgPlane
    Call GetFreePlane(tmpBayerPlane, "allbayer", idpDepthS16, False, "Clamp Image (Bayer plane)")
    Call Copy(InputPlane, InputZone, EEE_COLOR_FLAT, tmpBayerPlane, "ALLBAYER_ZONE2D", EEE_COLOR_FLAT)
    
    Call StdCM_SeparateRGBforBayer(tmpBayerPlane, "ALLBAYER_ZONE2D", _
                                   bayerRedPlane, "RBAYER_FULL", _
                                   bayerGreenPlane, "GBAYER_FULL", _
                                   bayerBluePlane, "BBAYER_FULL")

   
    Dim rRawPlane As CImgPlane
    Dim gRawPlane As CImgPlane
    Dim bRawPlane As CImgPlane
    Call GetFreePlane(rRawPlane, "pyline", idpDepthF32, False, "Red Raw Image")
    Call GetFreePlane(gRawPlane, "pyline", idpDepthF32, False, "Green Raw Image")
    Call GetFreePlane(bRawPlane, "pyline", idpDepthF32, False, "Blue Raw Image")
    Call Copy(bayerRedPlane, "RBAYER_FULL", EEE_COLOR_FLAT, rRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    Call Copy(bayerGreenPlane, "GBAYER_FULL", EEE_COLOR_FLAT, gRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    Call Copy(bayerBluePlane, "BBAYER_FULL", EEE_COLOR_FLAT, bRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)

    Call ReleasePlane(bayerRedPlane)
    Call ReleasePlane(bayerGreenPlane)
    Call ReleasePlane(bayerBluePlane)
    
    Dim redPlane As CImgPlane
    Dim greenPlane As CImgPlane
    Dim bluePlane As CImgPlane
    Call GetFreePlane(redPlane, "pyline", idpDepthF32, False, "red plane")
    Call GetFreePlane(greenPlane, "pyline", idpDepthF32, False, "green plane")
    Call GetFreePlane(bluePlane, "pyline", idpDepthF32, False, "blue plane")
    If IsClsMuraFlatFieldingOn Then
        Dim frShadingPlane As CImgPlane
        Dim fgShadingPlane As CImgPlane
        Dim fbShadingPlane As CImgPlane
        Set frShadingPlane = TheIDP.PlaneBank.Item("FLAT FIELD RED")
        Set fgShadingPlane = TheIDP.PlaneBank.Item("FLAT FIELD GREEN")
        Set fbShadingPlane = TheIDP.PlaneBank.Item("FLAT FIELD BLUE")
    
        Call Divide(rRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, _
                    frShadingPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, _
                    redPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
        Call Divide(gRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, _
                    fgShadingPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, _
                    greenPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
        Call Divide(bRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, _
                    fbShadingPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, _
                    bluePlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)

        Call ReleasePlane(frShadingPlane)
        Call ReleasePlane(fgShadingPlane)
        Call ReleasePlane(fbShadingPlane)
    Else
        Call Copy(rRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, redPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
        Call Copy(gRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, greenPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
        Call Copy(bRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, bluePlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    End If

    Call ReleasePlane(rRawPlane)
    Call ReleasePlane(gRawPlane)
    Call ReleasePlane(bRawPlane)
    
    '以降の処理で必要なので、Ｒ，Ｇ，Ｂの画像をプレーンバンクに登録
    Call TheIDP.PlaneBank.Add("MURA_RED_SEP_PLANE", redPlane, , True)
    Call TheIDP.PlaneBank.Add("MURA_GREEN_SEP_PLANE", greenPlane, , True)
    Call TheIDP.PlaneBank.Add("MURA_BLUE_SEP_PLANE", bluePlane, , True)
    
    '========== MEAN VALUES OF R/G/B and Y ==================================
    Dim redMean(nSite) As Double
    Dim greenMean(nSite) As Double
    Dim blueMean(nSite) As Double
    Call Average(redPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, redMean)
    Call Average(greenPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, greenMean)
    Call Average(bluePlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, blueMean)
    
    For site = 0 To nSite
        result(rgbColorArray.red, site) = redMean(site)
        result(rgbColorArray.green, site) = greenMean(site)
        result(rgbColorArray.blue, site) = blueMean(site)
    Next site

End Sub

Public Sub CM_ycode(ByRef rgbMean() As Double, ByRef result() As Double)

    Dim site As Long

    Dim tmpFFMeanR() As Double
    Dim tmpFFMeanG() As Double
    Dim tmpFFMeanB() As Double

    Call stdCM_GetFFMean(tmpFFMeanR, tmpFFMeanG, tmpFFMeanB)
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            result(site) = rgbMean(rgbColorArray.red, site) * 0.3 * tmpFFMeanR(site) _
                         + rgbMean(rgbColorArray.green, site) * 0.59 * tmpFFMeanG(site) _
                         + rgbMean(rgbColorArray.blue, site) * 0.11 * tmpFFMeanB(site)
        End If
    Next site

End Sub

Public Sub CM_wbalance_lpf(ByRef rgbMean() As Double, ByRef yLevel() As Double)

    Dim site As Long

    Dim redPlane As CImgPlane, greenPlane As CImgPlane, bluePlane As CImgPlane
    
    Call GetRegisteredPlane("MURA_RED_SEP_PLANE", redPlane, "redPlane", True)       'Used Plane Delete
    Call GetRegisteredPlane("MURA_GREEN_SEP_PLANE", greenPlane, "greenPlane", True) 'Used Plane Delete
    Call GetRegisteredPlane("MURA_BLUE_SEP_PLANE", bluePlane, "bluePlane", True)    'Used Plane Delete

        '========== WB parameters ==================================
    Dim yFactorR(nSite) As Double
    Dim yFactorG(nSite) As Double
    Dim yFactorB(nSite) As Double
    Dim rRegisterMean(nSite) As Double
    Dim gRegisterMean(nSite) As Double
    Dim bRegisterMean(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            If rgbMean(rgbColorArray.red, site) <> 0# Then
                yFactorR(site) = 0.3 / rgbMean(rgbColorArray.red, site)
                rRegisterMean(site) = 0.3
            Else
                yFactorR(site) = 0#
                rRegisterMean(site) = 0#
            End If
            
            If rgbMean(rgbColorArray.green, site) <> 0# Then
                yFactorG(site) = 0.59 / rgbMean(rgbColorArray.green, site)
                gRegisterMean(site) = 0.59
            Else
                yFactorG(site) = 0#
                gRegisterMean(site) = 0#
            End If
                
            If rgbMean(rgbColorArray.blue, site) <> 0# Then
                yFactorB(site) = 0.11 / rgbMean(rgbColorArray.blue, site)
                bRegisterMean(site) = 0.11
            Else
                yFactorB(site) = 0#
                bRegisterMean(site) = 0#
            End If
        End If
    Next site
    
    '========== LPF ==================================
    Dim fWorkPlane0 As CImgPlane
    Call GetFreePlane(fWorkPlane0, redPlane.planeGroup, idpDepthF32, False, "Work Plane 0")
    Dim fLowPassRed As CImgPlane
    Dim fLowPassGreen As CImgPlane
    Dim fLowPassBlue As CImgPlane
    
    Call GetFreePlane(fLowPassRed, redPlane.planeGroup, idpDepthF32, False, "Low Pass Red")
    Call MultiplyConst(redPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, yFactorR, fLowPassRed, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    '[IDV Point] "Low Pass Red"=(WB Red)*0.3 Before LPF
    Call StdCM_ApplyLowPassFilterYLine(fLowPassRed, fWorkPlane0, Kernel_LowPassH, Kernel_LowPassV)
    '[IDV Point] "Low Pass Red"=(WB Red)*0.3 After LPF
    
    Call GetFreePlane(fLowPassGreen, greenPlane.planeGroup, idpDepthF32, False, "Low Pass Green")
    Call MultiplyConst(greenPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, yFactorG, fLowPassGreen, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    '[IDV Point] "Low Pass Green"=(WB Green)*0.59 Before LPF
    Call StdCM_ApplyLowPassFilterYLine(fLowPassGreen, fWorkPlane0, Kernel_LowPassH, Kernel_LowPassV)
    '[IDV Point] "Low Pass Green"=(WB Green)*0.59 After LPF
    
    Call GetFreePlane(fLowPassBlue, bluePlane.planeGroup, idpDepthF32, False, "Low Pass Blue")
    Call MultiplyConst(bluePlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, yFactorB, fLowPassBlue, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    '[IDV Point] "Low Pass Blue"=(WB Blue)*0.11 Before LPF
    Call StdCM_ApplyLowPassFilterYLine(fLowPassBlue, fWorkPlane0, Kernel_LowPassH, Kernel_LowPassV)
    '[IDV Point] "Low Pass Blue"=(WB Blue)*0.11 After LPF

'========================================================================================================
'''To view white balance R/G/B images, uncommnet the following line (otherwise, comment it).
'Call StdCM_WBImage(fLowPassRed, fLowPassGreen, fLowPassBlue)
'========================================================================================================

    '========== To GENERATE Y IMAGE ==================================
    Dim yPlane As CImgPlane
    Call GetFreePlane(yPlane, fLowPassRed.planeGroup, idpDepthF32, False, "Y Plane")
    Call Add(fLowPassRed, "YLINE_ZONE2D", EEE_COLOR_FLAT, fLowPassGreen, "YLINE_ZONE2D", EEE_COLOR_FLAT, fWorkPlane0, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    Call Add(fWorkPlane0, "YLINE_ZONE2D", EEE_COLOR_FLAT, fLowPassBlue, "YLINE_ZONE2D", EEE_COLOR_FLAT, yPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    '[IDV Point] "Y Plane"=Y

    '以降の処理で必要なので、Ｙ, Ｒ，Ｇ，Ｂの画像をプレーンバンクに登録
    Call TheIDP.PlaneBank.Add("MURA_LP_YLINE_PLANE", yPlane, , True)
    Call TheIDP.PlaneBank.Add("MURA_LP_RED_WB_PLANE", fLowPassRed, , True)
    Call TheIDP.PlaneBank.Add("MURA_LP_GREEN_WB_PLANE", fLowPassGreen, , True)
    Call TheIDP.PlaneBank.Add("MURA_LP_BLUE_WB_PLANE", fLowPassBlue, , True)
    
     '========== Y_MEAN ==================================
    Call Average(yPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, yLevel)
    Call StdCM_RegisteredValue("Y_MEAN", yLevel)
    
    Call StdCM_RegisteredValue("R_MEAN", rRegisterMean)
    Call StdCM_RegisteredValue("G_MEAN", gRegisterMean)
    Call StdCM_RegisteredValue("B_MEAN", bRegisterMean)

End Sub

Public Sub CMT_wbalance_lpf(ByRef rgbMean() As Double, ByRef yLevel() As Double)    '2繋ぎ

    Dim site As Long

    Dim redPlane As CImgPlane, greenPlane As CImgPlane, bluePlane As CImgPlane
    
    Call GetRegisteredPlane("MURA_RED_SEP_PLANE", redPlane, "redPlane", True)       'Used Plane Delete
    Call GetRegisteredPlane("MURA_GREEN_SEP_PLANE", greenPlane, "greenPlane", True) 'Used Plane Delete
    Call GetRegisteredPlane("MURA_BLUE_SEP_PLANE", bluePlane, "bluePlane", True)    'Used Plane Delete

        '========== WB parameters ==================================
    Dim yFactorR(nSite) As Double
    Dim yFactorG(nSite) As Double
    Dim yFactorB(nSite) As Double
    Dim rRegisterMean(nSite) As Double
    Dim gRegisterMean(nSite) As Double
    Dim bRegisterMean(nSite) As Double
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            If rgbMean(rgbColorArray.red, site) <> 0# Then
                yFactorR(site) = 0.3 / rgbMean(rgbColorArray.red, site)
                rRegisterMean(site) = 0.3
            Else
                yFactorR(site) = 0#
                rRegisterMean(site) = 0#
            End If
            
            If rgbMean(rgbColorArray.green, site) <> 0# Then
                yFactorG(site) = 0.59 / rgbMean(rgbColorArray.green, site)
                gRegisterMean(site) = 0.59
            Else
                yFactorG(site) = 0#
                gRegisterMean(site) = 0#
            End If
                
            If rgbMean(rgbColorArray.blue, site) <> 0# Then
                yFactorB(site) = 0.11 / rgbMean(rgbColorArray.blue, site)
                bRegisterMean(site) = 0.11
            Else
                yFactorB(site) = 0#
                bRegisterMean(site) = 0#
            End If
        End If
    Next site
    
    '========== LPF ==================================
    Dim fWorkPlane0 As CImgPlane
    Call GetFreePlane(fWorkPlane0, redPlane.planeGroup, idpDepthF32, False, "Work Plane 0")
    Dim fLowPassRed As CImgPlane
    Dim fLowPassGreen As CImgPlane
    Dim fLowPassBlue As CImgPlane
    
    Call GetFreePlane(fLowPassRed, redPlane.planeGroup, idpDepthF32, False, "Low Pass Red")
    Call MultiplyConst(redPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, yFactorR, fLowPassRed, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    '[IDV Point] "Low Pass Red"=(WB Red)*0.3 Before LPF
    Call StdCMT_ApplyLowPassFilterYLine(fLowPassRed, fWorkPlane0, Kernel_LowPassH, Kernel_LowPassV)
    '[IDV Point] "Low Pass Red"=(WB Red)*0.3 After LPF
    
    Call GetFreePlane(fLowPassGreen, greenPlane.planeGroup, idpDepthF32, False, "Low Pass Green")
    Call MultiplyConst(greenPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, yFactorG, fLowPassGreen, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    '[IDV Point] "Low Pass Green"=(WB Green)*0.59 Before LPF
    Call StdCMT_ApplyLowPassFilterYLine(fLowPassGreen, fWorkPlane0, Kernel_LowPassH, Kernel_LowPassV)
    '[IDV Point] "Low Pass Green"=(WB Green)*0.59 After LPF
    
    Call GetFreePlane(fLowPassBlue, bluePlane.planeGroup, idpDepthF32, False, "Low Pass Blue")
    Call MultiplyConst(bluePlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, yFactorB, fLowPassBlue, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    '[IDV Point] "Low Pass Blue"=(WB Blue)*0.11 Before LPF
    Call StdCMT_ApplyLowPassFilterYLine(fLowPassBlue, fWorkPlane0, Kernel_LowPassH, Kernel_LowPassV)
    '[IDV Point] "Low Pass Blue"=(WB Blue)*0.11 After LPF

'========================================================================================================
'''To view white balance R/G/B images, uncommnet the following line (otherwise, comment it).
'Call StdCM_WBImage(fLowPassRed, fLowPassGreen, fLowPassBlue)
'========================================================================================================

    '========== To GENERATE Y IMAGE ==================================
    Dim yPlane As CImgPlane
    Call GetFreePlane(yPlane, fLowPassRed.planeGroup, idpDepthF32, False, "Y Plane")
    Call Add(fLowPassRed, "YLINE_ZONE2D", EEE_COLOR_FLAT, fLowPassGreen, "YLINE_ZONE2D", EEE_COLOR_FLAT, fWorkPlane0, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    Call Add(fWorkPlane0, "YLINE_ZONE2D", EEE_COLOR_FLAT, fLowPassBlue, "YLINE_ZONE2D", EEE_COLOR_FLAT, yPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    '[IDV Point] "Y Plane"=Y

    '以降の処理で必要なので、Ｙ, Ｒ，Ｇ，Ｂの画像をプレーンバンクに登録
    Call TheIDP.PlaneBank.Add("MURA_LP_YLINE_PLANE", yPlane, , True)
    Call TheIDP.PlaneBank.Add("MURA_LP_RED_WB_PLANE", fLowPassRed, , True)
    Call TheIDP.PlaneBank.Add("MURA_LP_GREEN_WB_PLANE", fLowPassGreen, , True)
    Call TheIDP.PlaneBank.Add("MURA_LP_BLUE_WB_PLANE", fLowPassBlue, , True)
    
     '========== Y_MEAN ==================================
    Call Average(yPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, yLevel)
    Call StdCM_RegisteredValue("Y_MEAN", yLevel)
    
    Call StdCM_RegisteredValue("R_MEAN", rRegisterMean)
    Call StdCM_RegisteredValue("G_MEAN", gRegisterMean)
    Call StdCM_RegisteredValue("B_MEAN", bRegisterMean)

End Sub

Public Sub CM_yline(Slice As Variant, ByRef result() As Double)

    Dim yPlane As CImgPlane
    
    Call GetRegisteredPlane("MURA_LP_YLINE_PLANE", yPlane, "yPlane")

    Call std_CalcYline(yPlane, Slice, yLineDif, yLineCoef, result)

End Sub

Public Sub CMT_yline(Slice As Variant, ByRef result() As Double)

    Dim yPlane As CImgPlane
    
    Call GetRegisteredPlane("MURA_LP_YLINE_PLANE", yPlane, "yPlane")

    Call std_TCalcYline(yPlane, Slice, yLineDif, yLineCoef, result)

End Sub

Public Sub CM_ResizeYlineYlocal()

    Dim yPlane As CImgPlane
    
    Call GetRegisteredPlane("MURA_LP_YLINE_PLANE", yPlane, "yPlane")

    Dim localPlane As CImgPlane
    Call GetFreePlane(localPlane, "pylocal", idpDepthF32, False, "local source")
    Call MultiMean(yPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, localPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, COMP_YLINE_TO_YLOCAL, COMP_YLINE_TO_YLOCAL)  'localPlane -> use Y_global

    Call TheIDP.PlaneBank.Add("MURA_ReSize_YLOCAL_PLANE", localPlane, , True)

End Sub

Public Sub CM_yLocalDiff()

    Dim localPlane As CImgPlane
    
    Call GetRegisteredPlane("MURA_ReSize_YLOCAL_PLANE", localPlane, "ylocalPlane")
    
    Dim hDifPlane As CImgPlane, vDifPlane As CImgPlane
    Call GetFreePlane(hDifPlane, localPlane.planeGroup, idpDepthF32, , "hDifPlane")
    Call GetFreePlane(vDifPlane, localPlane.planeGroup, idpDepthF32, , "vDifPlane")
    
    Dim workPlane1 As CImgPlane
    Call GetFreePlane(workPlane1, localPlane.planeGroup, idpDepthF32)
    Call Copy(localPlane, "YLOCAL_COL_SUB_TARGET", EEE_COLOR_FLAT, workPlane1, "YLOCAL_COL_SUB_TARGET", EEE_COLOR_FLAT)
    Call Subtract(localPlane, "YLOCAL_COL_SUB_SOURCE", EEE_COLOR_FLAT, _
                  workPlane1, "YLOCAL_COL_SUB_TARGET", EEE_COLOR_FLAT, _
                  hDifPlane, "YLOCAL_COL_SUB_SOURCE", EEE_COLOR_FLAT)   ' hDifPlane -> use Yfram2D
    
    Call SubRows(localPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, vDifPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, yLocalDif)   ' vDifPlane -> use Yfram2D

    Call TheIDP.PlaneBank.Add("MURA_YLOCAL_HDIFF_PLANE", hDifPlane)
    Call TheIDP.PlaneBank.Add("MURA_YLOCAL_VDIFF_PLANE", vDifPlane)

End Sub

Public Sub CMT_yLocalDiff()

    Dim localPlane As CImgPlane
    
    Call GetRegisteredPlane("MURA_ReSize_YLOCAL_PLANE", localPlane, "ylocalPlane")
    
     '========== H_DIFF TSUNAGI ==================================
    Dim hDifPlane As CImgPlane, vDifPlane As CImgPlane
    Call GetFreePlane(hDifPlane, localPlane.planeGroup, idpDepthF32, , "hDifPlane")
    Call GetFreePlane(vDifPlane, localPlane.planeGroup, idpDepthF32, , "vDifPlane")
    
    Dim workPlane1 As CImgPlane
    Call GetFreePlane(workPlane1, localPlane.planeGroup, idpDepthF32)
    Call Copy(localPlane, "YLOCAL_COL_SUB_TARGET_L", EEE_COLOR_FLAT, workPlane1, "YLOCAL_COL_SUB_TARGET_L", EEE_COLOR_FLAT)
    Call Subtract(localPlane, "YLOCAL_COL_SUB_SOURCE_L", EEE_COLOR_FLAT, _
                  workPlane1, "YLOCAL_COL_SUB_TARGET_L", EEE_COLOR_FLAT, _
                  hDifPlane, "YLOCAL_COL_SUB_SOURCE_L", EEE_COLOR_FLAT)   ' hDifPlane -> use Yfram2D

    Dim workPlane2 As CImgPlane
    Call GetFreePlane(workPlane2, localPlane.planeGroup, idpDepthF32)
    Call Copy(localPlane, "YLOCAL_COL_SUB_TARGET_R", EEE_COLOR_FLAT, workPlane2, "YLOCAL_COL_SUB_TARGET_R", EEE_COLOR_FLAT)
    Call Subtract(localPlane, "YLOCAL_COL_SUB_SOURCE_R", EEE_COLOR_FLAT, _
                  workPlane2, "YLOCAL_COL_SUB_TARGET_R", EEE_COLOR_FLAT, _
                  hDifPlane, "YLOCAL_COL_SUB_SOURCE_R", EEE_COLOR_FLAT)   ' hDifPlane -> use Yfram2D
    
     '========== V_DIFF ==================================
    Call SubRows(localPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, vDifPlane, "YLOCAL_ZONE2D", EEE_COLOR_FLAT, yLocalDif)   ' vDifPlane -> use Yfram2D

    Call TheIDP.PlaneBank.Add("MURA_YLOCAL_HDIFF_PLANE", hDifPlane)
    Call TheIDP.PlaneBank.Add("MURA_YLOCAL_VDIFF_PLANE", vDifPlane)

End Sub

Public Sub CM_ylocal(Slice As Variant, ByRef result() As Double)

    Dim site As Long
    Dim tmpYlocal(nSite) As Double
    Dim tmpYlUpp(nSite) As Double

    Dim hDifPlane As CImgPlane, vDifPlane As CImgPlane
    
    Call GetRegisteredPlane("MURA_YLOCAL_HDIFF_PLANE", hDifPlane, "hdifPlane")
    Call GetRegisteredPlane("MURA_YLOCAL_VDIFF_PLANE", vDifPlane, "vdifPlane")

    Dim workPlane_h As CImgPlane
    Dim workPlane_v As CImgPlane
    Call GetFreePlane(workPlane_h, hDifPlane.planeGroup, idpDepthF32)
    Call GetFreePlane(workPlane_v, hDifPlane.planeGroup, idpDepthF32)
    Call Copy(hDifPlane, "YLOCAL_FULL", EEE_COLOR_FLAT, workPlane_h, "YLOCAL_FULL", EEE_COLOR_FLAT)
    Call Copy(vDifPlane, "YLOCAL_FULL", EEE_COLOR_FLAT, workPlane_v, "YLOCAL_FULL", EEE_COLOR_FLAT)

    Call std_CalcYlocal(workPlane_h, workPlane_v, Slice, yLocalDif, yLocalCoef, tmpYlocal, tmpYlUpp)

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            result(CMReturnType.rNUM, site) = tmpYlocal(site)
            result(CMReturnType.rMAX, site) = tmpYlUpp(site)
        End If
    Next site

End Sub

Public Sub CM_yfrm2d(Slice As Variant, ByRef result() As Double)

    Dim site As Long
    Dim ret_tmp1(nSite) As Double
    Dim ret_tmp2(nSite) As Double
    
    Dim hDifPlane As CImgPlane, vDifPlane As CImgPlane
    
    Call GetRegisteredPlane("MURA_YLOCAL_HDIFF_PLANE", hDifPlane, "hdifPlane")
    Call GetRegisteredPlane("MURA_YLOCAL_VDIFF_PLANE", vDifPlane, "vdifPlane")

    Dim workPlane_h As CImgPlane
    Dim workPlane_v As CImgPlane
    Call GetFreePlane(workPlane_h, hDifPlane.planeGroup, idpDepthF32)
    Call GetFreePlane(workPlane_v, hDifPlane.planeGroup, idpDepthF32)
    Call Copy(hDifPlane, "YLOCAL_FULL", EEE_COLOR_FLAT, workPlane_h, "YLOCAL_FULL", EEE_COLOR_FLAT)
    Call Copy(vDifPlane, "YLOCAL_FULL", EEE_COLOR_FLAT, workPlane_v, "YLOCAL_FULL", EEE_COLOR_FLAT)

    '========== ZONE1 DATA CLEAR ==========================
    Call WritePixel(workPlane_h, "YLOCAL_ZONE1", EEE_COLOR_FLAT, 0)
    Call WritePixel(workPlane_v, "YLOCAL_ZONE1", EEE_COLOR_FLAT, 0)

    Call std_CalcYfrm2d(workPlane_h, workPlane_v, Slice, yLocalDif, yLocalCoef, ret_tmp1, ret_tmp2)

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            result(CMReturnType.rNUM, site) = ret_tmp1(site)
            result(CMReturnType.rMAX, site) = ret_tmp2(site)
        End If
    Next site

End Sub

Public Sub CM_yfrm2(Slice As Variant, ByRef result() As Double)

    Dim site As Long
    Dim ret_tmp1(nSite) As Double
    Dim ret_tmp2(nSite) As Double
    
    Dim hDifPlane As CImgPlane, vDifPlane As CImgPlane
    
    Call GetRegisteredPlane("MURA_YLOCAL_HDIFF_PLANE", hDifPlane, "hdifPlane")
    Call GetRegisteredPlane("MURA_YLOCAL_VDIFF_PLANE", vDifPlane, "vdifPlane")

    Dim workPlane_h As CImgPlane
    Dim workPlane_v As CImgPlane
    Call GetFreePlane(workPlane_h, hDifPlane.planeGroup, idpDepthF32)
    Call GetFreePlane(workPlane_v, hDifPlane.planeGroup, idpDepthF32)
    Call Copy(hDifPlane, "YLOCAL_FULL", EEE_COLOR_FLAT, workPlane_h, "YLOCAL_FULL", EEE_COLOR_FLAT)
    Call Copy(vDifPlane, "YLOCAL_FULL", EEE_COLOR_FLAT, workPlane_v, "YLOCAL_FULL", EEE_COLOR_FLAT)

    '========== ZONE1 DATA CLEAR ==========================
    Call WritePixel(workPlane_h, "YLOCAL_ZONE1", EEE_COLOR_FLAT, 0)
    Call WritePixel(workPlane_v, "YLOCAL_ZONE1", EEE_COLOR_FLAT, 0)
    
    Call std_CalcYfrm2(workPlane_h, workPlane_v, Slice, yLocalDif, yLocalCoef, ret_tmp1, ret_tmp2)

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            result(CMReturnType.rNUM, site) = ret_tmp1(site)
            result(CMReturnType.rMAX, site) = ret_tmp2(site)
        End If
    Next site
    
End Sub

Public Sub CM_yglobal(Slice As Variant, ByRef result() As Double)

    Dim site As Long
    Dim ret_tmp1(nSite) As Double
    Dim ret_tmp2(nSite) As Double
    Dim localPlane As CImgPlane
    
    Call GetRegisteredPlane("MURA_ReSize_YLOCAL_PLANE", localPlane, "ylocalPlane")

    Dim workplane0 As CImgPlane
    Call GetFreePlane(workplane0, localPlane.planeGroup, idpDepthF32)
    Call Copy(localPlane, "YLOCAL_FULL", EEE_COLOR_FLAT, workplane0, "YLOCAL_FULL", EEE_COLOR_FLAT)

    Call std_CalcYglob(workplane0, Slice, yGlobDif, yGlobCoef, ret_tmp1, ret_tmp2)

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            result(CMReturnType.rNUM, site) = ret_tmp1(site)
            result(CMReturnType.rMAX, site) = ret_tmp2(site)
        End If
    Next site

End Sub

Public Sub CMT_yglobal(Slice As Variant, ByRef result() As Double)

    Dim site As Long
    Dim ret_tmp1(nSite) As Double
    Dim ret_tmp2(nSite) As Double
    Dim localPlane As CImgPlane
    
    Call GetRegisteredPlane("MURA_ReSize_YLOCAL_PLANE", localPlane, "ylocalPlane")

    Dim workplane0 As CImgPlane
    Call GetFreePlane(workplane0, localPlane.planeGroup, idpDepthF32)
    Call Copy(localPlane, "YLOCAL_FULL", EEE_COLOR_FLAT, workplane0, "YLOCAL_FULL", EEE_COLOR_FLAT)

    Call std_TCalcYglob(workplane0, Slice, yGlobDif, yGlobCoef, ret_tmp1, ret_tmp2)

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            result(CMReturnType.rNUM, site) = ret_tmp1(site)
            result(CMReturnType.rMAX, site) = ret_tmp2(site)
        End If
    Next site

End Sub

Public Sub CM_yglobal2(Slice As Variant, ByRef result() As Double)

    Dim site As Long
    Dim ret_tmp1(nSite) As Double
    Dim ret_tmp2(nSite) As Double
    Dim yPlane As CImgPlane
    
    Call GetRegisteredPlane("MURA_LP_YLINE_PLANE", yPlane, "yPlane")

    Dim workplane0 As CImgPlane
    Call GetFreePlane(workplane0, yPlane.planeGroup, idpDepthF32)
    Call Copy(yPlane, "YLINE_FULL", EEE_COLOR_FLAT, workplane0, "YLINE_FULL", EEE_COLOR_FLAT)

    Call std_CalcYglob2(workplane0, Slice, yGlob2HDif, yGlob2VDif, ret_tmp1, ret_tmp2)

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            result(CMReturnType.rNUM, site) = ret_tmp1(site)
            result(CMReturnType.rMAX, site) = ret_tmp2(site)
        End If
    Next site

End Sub

Public Sub CMT_yglobal2(Slice As Variant, ByRef result() As Double)

    Dim site As Long
    Dim ret_tmp1(nSite) As Double
    Dim ret_tmp2(nSite) As Double
    Dim yPlane As CImgPlane
    
    Call GetRegisteredPlane("MURA_LP_YLINE_PLANE", yPlane, "yPlane")

    Dim workplane0 As CImgPlane
    Call GetFreePlane(workplane0, yPlane.planeGroup, idpDepthF32)
    Call Copy(yPlane, "YLINE_FULL", EEE_COLOR_FLAT, workplane0, "YLINE_FULL", EEE_COLOR_FLAT)

    Call std_TCalcYglob2(workplane0, Slice, yGlob2HDif, yGlob2VDif, ret_tmp1, ret_tmp2)

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            result(CMReturnType.rNUM, site) = ret_tmp1(site)
            result(CMReturnType.rMAX, site) = ret_tmp2(site)
        End If
    Next site

End Sub

Public Sub CM_makeChroma()

    Dim fLowPassRed As CImgPlane
    Dim fLowPassGreen As CImgPlane
    Dim fLowPassBlue As CImgPlane
    Call GetRegisteredPlane("MURA_LP_RED_WB_PLANE", fLowPassRed, "red Low Pass Plane")
    Call GetRegisteredPlane("MURA_LP_GREEN_WB_PLANE", fLowPassGreen, "green Low Pass Plane")
    Call GetRegisteredPlane("MURA_LP_BLUE_WB_PLANE", fLowPassBlue, "blue Low Pass Plane")
    
    Dim gainedRyPlane As CImgPlane
    Dim gainedByPlane As CImgPlane
    Call GetFreePlane(gainedRyPlane, "pclocal", idpDepthF32, , "gained Ry plane")
    Call GetFreePlane(gainedByPlane, "pclocal", idpDepthF32, , "gained By plane")

    Dim clocalWorkPlane0 As CImgPlane
    Dim clocalWorkPlane1 As CImgPlane
    Call GetFreePlane(clocalWorkPlane0, "pclocal", idpDepthF32, False, "clocal work plane0")
    Call GetFreePlane(clocalWorkPlane1, "pclocal", idpDepthF32, False, "clocal work plane1")

'Clocal圧縮～RY、BY生成まで
    Call pre_CalcClocal(fLowPassRed, fLowPassGreen, fLowPassBlue, gainedRyPlane, gainedByPlane)  'input : fLowPassRed,fLowPassGreen,fLowPassBlue  output :gainedRyPlane,gainedByPlane

    Call TheIDP.PlaneBank.Add("MURA_GAINED_R_Y_PLANE", gainedRyPlane)
    Call TheIDP.PlaneBank.Add("MURA_GAINED_B_Y_PLANE", gainedByPlane)

End Sub

Public Sub CM_cLocalDiff()

    Dim gainedRyPlane As CImgPlane
    Dim gainedByPlane As CImgPlane
    Call GetRegisteredPlane("MURA_GAINED_R_Y_PLANE", gainedRyPlane, "gained Ry plane")
    Call GetRegisteredPlane("MURA_GAINED_B_Y_PLANE", gainedByPlane, "gained By plane")

    Dim ryHDifPlane As CImgPlane
    Dim byHDifPlane As CImgPlane
    Dim ryVDifPlane As CImgPlane
    Dim byVDifPlane As CImgPlane
    Call GetFreePlane(ryHDifPlane, gainedRyPlane.planeGroup, idpDepthF32, True, "ryHDifPlane")
    Call GetFreePlane(ryVDifPlane, gainedRyPlane.planeGroup, idpDepthF32, True, "ryVDifPlane")
    Call GetFreePlane(byHDifPlane, gainedByPlane.planeGroup, idpDepthF32, True, "byHDifPlane")
    Call GetFreePlane(byVDifPlane, gainedByPlane.planeGroup, idpDepthF32, True, "byVDifPlane")

    Dim clocalWorkPlane0 As CImgPlane
    Call GetFreePlane(clocalWorkPlane0, gainedRyPlane.planeGroup, idpDepthF32, False, "clocal work plane0")
    '========== DIFF ======================================
    '-------- H-DIFF --------------------------------------
    Call Copy(gainedRyPlane, "CLOCAL_COL_SUB_TARGET", EEE_COLOR_FLAT, clocalWorkPlane0, "CLOCAL_COL_SUB_TARGET", EEE_COLOR_FLAT)
    Call Subtract(gainedRyPlane, "CLOCAL_COL_SUB_SOURCE", EEE_COLOR_FLAT, _
                  clocalWorkPlane0, "CLOCAL_COL_SUB_TARGET", EEE_COLOR_FLAT, _
                  ryHDifPlane, "CLOCAL_COL_SUB_SOURCE", EEE_COLOR_FLAT)
    
    Call Copy(gainedByPlane, "CLOCAL_COL_SUB_TARGET", EEE_COLOR_FLAT, clocalWorkPlane0, "CLOCAL_COL_SUB_TARGET", EEE_COLOR_FLAT)
    Call Subtract(gainedByPlane, "CLOCAL_COL_SUB_SOURCE", EEE_COLOR_FLAT, _
                  clocalWorkPlane0, "CLOCAL_COL_SUB_TARGET", EEE_COLOR_FLAT, _
                  byHDifPlane, "CLOCAL_COL_SUB_SOURCE", EEE_COLOR_FLAT)
    
    '-------- V-DIFF --------------------------------------
    Call SubRows(gainedRyPlane, "CLOCAL_ZONE2D", EEE_COLOR_FLAT, ryVDifPlane, "CLOCAL_ZONE2D", EEE_COLOR_FLAT, cLocalDif)
    Call SubRows(gainedByPlane, "CLOCAL_ZONE2D", EEE_COLOR_FLAT, byVDifPlane, "CLOCAL_ZONE2D", EEE_COLOR_FLAT, cLocalDif)
    '[IDV Point] "ryHDifPlane"= RY H-Diff
    '[IDV Point] "ryVDifPlane"= RY V-DIff
    '[IDV Point] "byHDifPlane"= BY H-Diff
    '[IDV Point] "byVDifPlane"= BY V-DIff
 
     Dim clocalWorkPlane1 As CImgPlane
     Call GetFreePlane(clocalWorkPlane1, ryHDifPlane.planeGroup, idpDepthF32, False, "clocal work plane1")
    'The derived square of euclidean distance on the Ry-By plane will be put in
    '"ryHDifPlane", since this plane will no longer be used.
    Call Multiply(ryHDifPlane, "CLOCAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, _
                  ryHDifPlane, "CLOCAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, _
                  clocalWorkPlane0, "CLOCAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT)
    Call Multiply(byHDifPlane, "CLOCAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, _
                  byHDifPlane, "CLOCAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, _
                  clocalWorkPlane1, "CLOCAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT)
    Call Add(clocalWorkPlane0, "CLOCAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, _
             clocalWorkPlane1, "CLOCAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, _
             ryHDifPlane, "CLOCAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT)

    Call Multiply(ryVDifPlane, "CLOCAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, _
                  ryVDifPlane, "CLOCAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, _
                  clocalWorkPlane0, "CLOCAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT)
    Call Multiply(byVDifPlane, "CLOCAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, _
                  byVDifPlane, "CLOCAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, _
                  clocalWorkPlane1, "CLOCAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT)
    Call Add(clocalWorkPlane0, "CLOCAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, _
             clocalWorkPlane1, "CLOCAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, _
             ryVDifPlane, "CLOCAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT)

    Call TheIDP.PlaneBank.Add("MURA_CHROMA_HDIFF_PLANE", ryHDifPlane)
    Call TheIDP.PlaneBank.Add("MURA_CHROMA_VDIFF_PLANE", ryVDifPlane)

End Sub

Public Sub CM_clocal(ByRef result() As Double)

    Dim ryHDifPlane As CImgPlane
    Dim ryVDifPlane As CImgPlane
    Call GetRegisteredPlane("MURA_CHROMA_HDIFF_PLANE", ryHDifPlane, "chroma h-diff plane")
    Call GetRegisteredPlane("MURA_CHROMA_VDIFF_PLANE", ryVDifPlane, "chroma v-diff plane")

    Dim workPlane_h As CImgPlane
    Dim workPlane_v As CImgPlane
    Call GetFreePlane(workPlane_h, ryHDifPlane.planeGroup, idpDepthF32)
    Call GetFreePlane(workPlane_v, ryHDifPlane.planeGroup, idpDepthF32)
    Call Copy(ryHDifPlane, "CLOCAL_FULL", EEE_COLOR_FLAT, workPlane_h, "CLOCAL_FULL", EEE_COLOR_FLAT)
    Call Copy(ryVDifPlane, "CLOCAL_FULL", EEE_COLOR_FLAT, workPlane_v, "CLOCAL_FULL", EEE_COLOR_FLAT)

    Call std_CalcClocal(workPlane_h, workPlane_v, result)

End Sub


Public Sub CM_cframe2(ByRef result() As Double)

    Dim ryHDifPlane As CImgPlane
    Dim ryVDifPlane As CImgPlane
    Call GetRegisteredPlane("MURA_CHROMA_HDIFF_PLANE", ryHDifPlane, "chroma h-diff plane")
    Call GetRegisteredPlane("MURA_CHROMA_VDIFF_PLANE", ryVDifPlane, "chroma v-diff plane")

    Dim workPlane_h As CImgPlane
    Dim workPlane_v As CImgPlane
    Call GetFreePlane(workPlane_h, ryHDifPlane.planeGroup, idpDepthF32)
    Call GetFreePlane(workPlane_v, ryHDifPlane.planeGroup, idpDepthF32)
    Call Copy(ryHDifPlane, "CLOCAL_FULL", EEE_COLOR_FLAT, workPlane_h, "CLOCAL_FULL", EEE_COLOR_FLAT)
    Call Copy(ryVDifPlane, "CLOCAL_FULL", EEE_COLOR_FLAT, workPlane_v, "CLOCAL_FULL", EEE_COLOR_FLAT)

    'ZONE1をゼロクリア。
    Call WritePixel(workPlane_h, "CLOCAL_ZONE1", EEE_COLOR_FLAT, 0)
    Call WritePixel(workPlane_v, "CLOCAL_ZONE1", EEE_COLOR_FLAT, 0)

    Call std_CalcCfrm2(workPlane_h, workPlane_v, result)

End Sub

Public Sub CM_cframe2D(ByRef result() As Double)

    Dim ryHDifPlane As CImgPlane
    Dim ryVDifPlane As CImgPlane
    Call GetRegisteredPlane("MURA_CHROMA_HDIFF_PLANE", ryHDifPlane, "chroma h-diff plane")
    Call GetRegisteredPlane("MURA_CHROMA_VDIFF_PLANE", ryVDifPlane, "chroma v-diff plane")

    Dim workPlane_h As CImgPlane
    Dim workPlane_v As CImgPlane
    Call GetFreePlane(workPlane_h, ryHDifPlane.planeGroup, idpDepthF32)
    Call GetFreePlane(workPlane_v, ryHDifPlane.planeGroup, idpDepthF32)
    Call Copy(ryHDifPlane, "CLOCAL_FULL", EEE_COLOR_FLAT, workPlane_h, "CLOCAL_FULL", EEE_COLOR_FLAT)
    Call Copy(ryVDifPlane, "CLOCAL_FULL", EEE_COLOR_FLAT, workPlane_v, "CLOCAL_FULL", EEE_COLOR_FLAT)

    'ZONE1をゼロクリア。
    Call WritePixel(workPlane_h, "CLOCAL_ZONE1", EEE_COLOR_FLAT, 0)
    Call WritePixel(workPlane_v, "CLOCAL_ZONE1", EEE_COLOR_FLAT, 0)

    Call std_CalcCfrm2d(workPlane_h, workPlane_v, result)

End Sub

Public Sub CM_cglobalComp()

    Dim gainedRyPlane As CImgPlane
    Dim gainedByPlane As CImgPlane
    Call GetRegisteredPlane("MURA_GAINED_R_Y_PLANE", gainedRyPlane, "gained Ry plane")
    Call GetRegisteredPlane("MURA_GAINED_B_Y_PLANE", gainedByPlane, "gained By plane")

    '画像圧縮。(圧縮先のゾーンが"CGZ2D"ではなく、"CGZDD"に変わります)
    Dim ryGainedPlane As CImgPlane
    Dim byGainedPlane As CImgPlane
    Call GetFreePlane(ryGainedPlane, "pcglobal", idpDepthF32, False, "ryGainedPlane")
    Call GetFreePlane(byGainedPlane, "pcglobal", idpDepthF32, False, "byGainedPlane")
    Call MultiMean(gainedRyPlane, "CLOCAL_ZONE2D", EEE_COLOR_FLAT, _
                   ryGainedPlane, "CGLOBAL_ZONE2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, COMP_CLOCAL_TO_CGLOBAL, COMP_CLOCAL_TO_CGLOBAL)
    Call MultiMean(gainedByPlane, "CLOCAL_ZONE2D", EEE_COLOR_FLAT, _
                   byGainedPlane, "CGLOBAL_ZONE2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, COMP_CLOCAL_TO_CGLOBAL, COMP_CLOCAL_TO_CGLOBAL)

    Call TheIDP.PlaneBank.Add("MURA_COMP_R_Y_PLANE", ryGainedPlane, , True)
    Call TheIDP.PlaneBank.Add("MURA_COMP_B_Y_PLANE", byGainedPlane, , True)

End Sub


Public Sub CM_cglobalDiff()

    Dim ryGainedPlane As CImgPlane
    Dim byGainedPlane As CImgPlane
    Call GetRegisteredPlane("MURA_COMP_R_Y_PLANE", ryGainedPlane, "gained comp Ry plane")
    Call GetRegisteredPlane("MURA_COMP_B_Y_PLANE", byGainedPlane, "gained comp By plane")

    Dim cglo_ryHDifPlane As CImgPlane
    Dim cglo_byHDifPlane As CImgPlane
    Dim cglo_ryVDifPlane As CImgPlane
    Dim cglo_byVDifPlane As CImgPlane
    Call GetFreePlane(cglo_ryHDifPlane, ryGainedPlane.planeGroup, idpDepthF32, True, "ryHDifPlane")
    Call GetFreePlane(cglo_ryVDifPlane, ryGainedPlane.planeGroup, idpDepthF32, True, "ryVDifPlane")
    Call GetFreePlane(cglo_byHDifPlane, byGainedPlane.planeGroup, idpDepthF32, True, "byHDifPlane")
    Call GetFreePlane(cglo_byVDifPlane, byGainedPlane.planeGroup, idpDepthF32, True, "byVDifPlane")

    Dim cGlobalWorkPlane0 As CImgPlane
    Call GetFreePlane(cGlobalWorkPlane0, ryGainedPlane.planeGroup, idpDepthF32, False, "cglobal work plane0")
    '========== DIFF ======================================
    '-------- H-DIFF --------------------------------------
    Call Copy(ryGainedPlane, "CGLOBAL_COL_SUB_TARGET", EEE_COLOR_FLAT, cGlobalWorkPlane0, "CGLOBAL_COL_SUB_TARGET", EEE_COLOR_FLAT)
    Call Subtract(ryGainedPlane, "CGLOBAL_COL_SUB_SOURCE", EEE_COLOR_FLAT, _
                  cGlobalWorkPlane0, "CGLOBAL_COL_SUB_TARGET", EEE_COLOR_FLAT, _
                  cglo_ryHDifPlane, "CGLOBAL_COL_SUB_SOURCE", EEE_COLOR_FLAT)
    Call Copy(byGainedPlane, "CGLOBAL_COL_SUB_TARGET", EEE_COLOR_FLAT, cGlobalWorkPlane0, "CGLOBAL_COL_SUB_TARGET", EEE_COLOR_FLAT)
    Call Subtract(byGainedPlane, "CGLOBAL_COL_SUB_SOURCE", EEE_COLOR_FLAT, _
                  cGlobalWorkPlane0, "CGLOBAL_COL_SUB_TARGET", EEE_COLOR_FLAT, _
                  cglo_byHDifPlane, "CGLOBAL_COL_SUB_SOURCE", EEE_COLOR_FLAT)

    '-------- V-DIFF --------------------------------------
    Call SubRows(ryGainedPlane, "CGLOBAL_ZONE2D", EEE_COLOR_FLAT, cglo_ryVDifPlane, "CGLOBAL_ZONE2D", EEE_COLOR_FLAT, cGlobDif)
    Call SubRows(byGainedPlane, "CGLOBAL_ZONE2D", EEE_COLOR_FLAT, cglo_byVDifPlane, "CGLOBAL_ZONE2D", EEE_COLOR_FLAT, cGlobDif)

    Dim cGlobalWorkPlane1 As CImgPlane
    Call GetFreePlane(cGlobalWorkPlane1, cglo_ryHDifPlane.planeGroup, idpDepthF32, False, "cglobal work plane1")
    'The derived square of euclidean distance on the Ry-By plane will be put in
    '"ryHDifPlane", since this plane will no longer be used.
    Call Multiply(cglo_ryHDifPlane, "CGLOBAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, _
                  cglo_ryHDifPlane, "CGLOBAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, _
                  cGlobalWorkPlane0, "CGLOBAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT)
    Call Multiply(cglo_byHDifPlane, "CGLOBAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, _
                  cglo_byHDifPlane, "CGLOBAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, _
                  cGlobalWorkPlane1, "CGLOBAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT)
    Call Add(cGlobalWorkPlane0, "CGLOBAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, _
             cGlobalWorkPlane1, "CGLOBAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT, _
             cglo_ryHDifPlane, "CGLOBAL_FRAME_COL_JUDGE", EEE_COLOR_FLAT)

    Call Multiply(cglo_ryVDifPlane, "CGLOBAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, _
                  cglo_ryVDifPlane, "CGLOBAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, _
                  cGlobalWorkPlane0, "CGLOBAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT)
    Call Multiply(cglo_byVDifPlane, "CGLOBAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, _
                  cglo_byVDifPlane, "CGLOBAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, _
                  cGlobalWorkPlane1, "CGLOBAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT)
    Call Add(cGlobalWorkPlane0, "CGLOBAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, _
             cGlobalWorkPlane1, "CGLOBAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT, _
             cglo_ryVDifPlane, "CGLOBAL_FRAME_ROW_JUDGE", EEE_COLOR_FLAT)

    Call TheIDP.PlaneBank.Add("MURA_GLOBAL_CHROMA_HDIF_PLANE", cglo_ryHDifPlane)
    Call TheIDP.PlaneBank.Add("MURA_GLOBAL_CHROMA_VDIF_PLANE", cglo_ryVDifPlane)

End Sub

Public Sub CM_cglobal1(ByRef result() As Double)

    Dim cglo_ryHDifPlane As CImgPlane
    Dim cglo_ryVDifPlane As CImgPlane
    Call GetRegisteredPlane("MURA_GLOBAL_CHROMA_HDIF_PLANE", cglo_ryHDifPlane, "global chroma Hdif plane")
    Call GetRegisteredPlane("MURA_GLOBAL_CHROMA_VDIF_PLANE", cglo_ryVDifPlane, "global chroma Vdif plane")
    
    Dim workPlane_h As CImgPlane
    Dim workPlane_v As CImgPlane
    Call GetFreePlane(workPlane_h, cglo_ryHDifPlane.planeGroup, idpDepthF32)
    Call GetFreePlane(workPlane_v, cglo_ryHDifPlane.planeGroup, idpDepthF32)
    Call Copy(cglo_ryHDifPlane, "CGLOBAL_FULL", EEE_COLOR_FLAT, workPlane_h, "CGLOBAL_FULL", EEE_COLOR_FLAT)
    Call Copy(cglo_ryVDifPlane, "CGLOBAL_FULL", EEE_COLOR_FLAT, workPlane_v, "CGLOBAL_FULL", EEE_COLOR_FLAT)

    Call std_CalcCglob01(workPlane_h, workPlane_v, result)

End Sub

Public Sub CM_cglobal2d(ByRef result() As Double)

    Dim cglo_ryHDifPlane As CImgPlane
    Dim cglo_ryVDifPlane As CImgPlane
    Call GetRegisteredPlane("MURA_GLOBAL_CHROMA_HDIF_PLANE", cglo_ryHDifPlane, "global chroma Hdif plane")
    Call GetRegisteredPlane("MURA_GLOBAL_CHROMA_VDIF_PLANE", cglo_ryVDifPlane, "global chroma Vdif plane")

    Dim workPlane_h As CImgPlane
    Dim workPlane_v As CImgPlane
    Call GetFreePlane(workPlane_h, cglo_ryHDifPlane.planeGroup, idpDepthF32)
    Call GetFreePlane(workPlane_v, cglo_ryHDifPlane.planeGroup, idpDepthF32)
    Call Copy(cglo_ryHDifPlane, "CGLOBAL_FULL", EEE_COLOR_FLAT, workPlane_h, "CGLOBAL_FULL", EEE_COLOR_FLAT)
    Call Copy(cglo_ryVDifPlane, "CGLOBAL_FULL", EEE_COLOR_FLAT, workPlane_v, "CGLOBAL_FULL", EEE_COLOR_FLAT)

    Call std_CalcCglob2d(workPlane_h, workPlane_v, result)

End Sub

Public Sub CM_cshad(ByRef result() As Double)

    Dim ryGainedPlane As CImgPlane
    Dim byGainedPlane As CImgPlane
    Call GetRegisteredPlane("MURA_COMP_R_Y_PLANE", ryGainedPlane, "gained comp Ry plane")
    Call GetRegisteredPlane("MURA_COMP_B_Y_PLANE", byGainedPlane, "gained comp By plane")

    Dim workPlane_h As CImgPlane
    Dim workPlane_v As CImgPlane
    Call GetFreePlane(workPlane_h, ryGainedPlane.planeGroup, idpDepthF32)
    Call GetFreePlane(workPlane_v, ryGainedPlane.planeGroup, idpDepthF32)
    Call Copy(ryGainedPlane, "CGLOBAL_FULL", EEE_COLOR_FLAT, workPlane_h, "CGLOBAL_FULL", EEE_COLOR_FLAT)
    Call Copy(byGainedPlane, "CGLOBAL_FULL", EEE_COLOR_FLAT, workPlane_v, "CGLOBAL_FULL", EEE_COLOR_FLAT)

    Call std_CalcCshad(workPlane_h, workPlane_v, result)

End Sub

'===========================================================

Public Sub NbtdCM_Initialize(ByVal pType As String, ByVal ZONE_FULL As Variant, ByVal ZONE_ZONE3 As Variant, ByVal clampZone As Variant)

    'Local variables
    Dim site As Long                'For site loop

    Dim Flg_Active(nSite) As Long
    
    If TheIDP.KernelManager.IsExist("kernel_nbtdInitAntiAliasH") = True Then Exit Sub
    
    'ALL SITE ACTIVE
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Flg_Active(site) = 1
        End If
    Next site
    TheExec.sites.SetAllActive (True)

    Call StdBHCM_GetMuraParameters
    
    'Flat Field Mode
    Call BH_MakeFlatFieldImagePlane(pType, ZONE_FULL, ZONE_ZONE3, clampZone)
    
    'To define kernels.
    Call ker_nbtd
    
    For site = 0 To nSite
        If Flg_Active(site) = 0 Then
            TheExec.sites.site(site).Active = False
        End If
    Next site
    
End Sub

Private Sub BH_MakeFlatFieldImagePlane(ByVal pType As String, ByVal ZONE_FULL As Variant, ByVal ZONE_ZONE3 As Variant, ByVal clampZone As Variant)
    Dim site As Long
    Dim Bclamp(nSite) As Double     'For opb clamp
    
    'To read flat field image.
    Dim flatPlane As CImgPlane
    Call GetFreePlane(flatPlane, pType, idpDepthS16, True, "Flat Field Input")
    
'    Dim myTesterName As String
'    myTesterName = ETP_office & Format(Sw_Node, "000")
'    Dim myFileName As String
'
'    For site = 0 To nSite
'        myFileName = "FlatField_" & myTesterName & "_Site" & Format(site, "0") & ".stb"
'        Call InPutImage(site, flatPlane, ZONE_FULL, _
'                        Nbtd_FlatFieldPath & "\" & myFileName)
'    Next site

    Const FF_FILE_PREFIX As String = "HL-"
    Dim myTesterName As String
    myTesterName = Format(Sw_Node, "000")
    
    Dim myTesterSite As String
    Dim myFileName As String
    
    For site = 0 To nSite
        myTesterSite = "-Site-" & Format(site, "0")
        myFileName = FF_FILE_PREFIX & myTesterName & myTesterSite & ".stb"
        Call InPutImage(site, flatPlane, ZONE_FULL, Nbtd_FlatFieldPath & myTesterName & "\" & myFileName)
    Next site

    
    'OPB Clamp for Flat Field image.
    Call Average(flatPlane, clampZone, EEE_COLOR_FLAT, Bclamp)
    Dim workPlane1 As CImgPlane
    Call GetFreePlane(workPlane1, flatPlane.planeGroup, idpDepthS16, , "workPlane1")
    Call SubtractConst(flatPlane, ZONE_ZONE3, EEE_COLOR_FLAT, Bclamp, workPlane1, ZONE_ZONE3, EEE_COLOR_FLAT)

    'To perform noise reduction with median filter for flat field image.
    Dim workPlane2 As CImgPlane
    Call GetFreePlane(workPlane2, flatPlane.planeGroup, idpDepthS16, False, "workPlane2")
    Call Median(workPlane1, ZONE_ZONE3, idpColorAll, workPlane2, ZONE_ZONE3, idpColorAll, 5, 1) 'jikken
    Call Median(workPlane2, ZONE_ZONE3, idpColorAll, workPlane1, ZONE_ZONE3, idpColorAll, 1, 5) 'jikken
    Call ReleasePlane(workPlane2)

    'Bayer separation for flat field image.
    Dim bayerRedPlane As CImgPlane
    Dim bayerGreenPlane As CImgPlane
    Dim bayerBluePlane As CImgPlane
    Call GetFreePlane(bayerRedPlane, "rbayer", idpDepthS16, False, "redPlane for bayer")
    Call GetFreePlane(bayerGreenPlane, "gbayer", idpDepthS16, False, "greenPlane for bayer")
    Call GetFreePlane(bayerBluePlane, "bbayer", idpDepthS16, False, "bluePlane for bayer")
    
    Dim tmpBayerPlane As CImgPlane
    Call GetFreePlane(tmpBayerPlane, "allbayer", idpDepthS16, False, "Clamp Image (Bayer plane)")
    Call Copy(workPlane1, "ZONE2D_BAYER", EEE_COLOR_FLAT, tmpBayerPlane, "ALLBAYER_ZONE2D", EEE_COLOR_FLAT)
    Call StdCM_SeparateRGBforBayer(tmpBayerPlane, "ALLBAYER_ZONE2D", _
                                   bayerRedPlane, "RBAYER_FULL", _
                                   bayerGreenPlane, "GBAYER_FULL", _
                                   bayerBluePlane, "BBAYER_FULL")

    Call ReleasePlane(tmpBayerPlane)
    Call ReleasePlane(workPlane1)

    'To copy R/G/B flat field images to "YLINE" planes
    Dim bRawPlane As CImgPlane
    Call GetFreePlane(bRawPlane, "pyline", idpDepthF32, False, "Blue Flat Field")
    Call Copy(bayerBluePlane, "BBAYER_FULL", EEE_COLOR_FLAT, bRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    Call ReleasePlane(bayerBluePlane)
    
    'To deposit R/G/B flat field images to the Plane Bank.
    Call TheIDP.PlaneBank.Add("FLAT FIELD BLUE", bRawPlane, True, True)
    
    Call ReleasePlane(flatPlane)
    
End Sub

Public Sub ker_nbtd()

    With TheIDP
        'ベイヤー分割直後のBlue画像に対してかける圧縮エイリアシング防止用LPF(実は、1MHzのLPFと同じ係数でしたが、まずは別定義します)
        .CreateKernel "kernel_nbtdInitAntiAliasH", idpKernelInteger, 15, 1, 8, "-6 -4 0 9 22 36 46 50 46 36 22 9 0 -4 -6" '256
        .CreateKernel "kernel_nbtdInitAntiAliasV", idpKernelInteger, 1, 15, 8, "-6 -4 0 9 22 36 46 50 46 36 22 9 0 -4 -6"
               
        '---- NBSJTP ----
        '微分時の周波数選択用
        .CreateKernel "kernel_nbtdFreqSelectH", idpKernelInteger, 15, 1, 8, "5 -6 -13 -13 5 38 70 84 70 38 5 -13 -13 -6 5" '256
        .CreateKernel "kernel_nbtdFreqSelectV", idpKernelInteger, 1, 15, 8, "5 -6 -13 -13 5 38 70 84 70 38 5 -13 -13 -6 5"
        
'        'シェーディング除去用。変更する場合に"Nbtd_ShadingFilterTap"の変更も忘れずに
        .CreateKernel "kernel_nbtdRmShadingH", idpKernelInteger, 11, 1, 0, "1 1 1 1 1 1 1 1 1 1 1"
        .CreateKernel "kernel_nbtdRmShadingV", idpKernelInteger, 1, 11, 0, "1 1 1 1 1 1 1 1 1 1 1"
        
    End With
    
End Sub

Public Sub Pre_BHakimura(ByRef inBluePlane As CImgPlane, ByRef bluePlane As CImgPlane, ByRef sjMean() As Double, ByRef multiplyFactors() As Double, ByRef hDifPlane As CImgPlane, ByRef vDifPlane As CImgPlane)

'    Const ShadingFilterTap As Long = BHShadingFilterTap
 
    Dim expSize As Long
    Dim rmShadFilterH As String
    Dim rmShadFilterV As String
    expSize = Int(BHShadingFilterTap / 2)
    rmShadFilterH = "kernel_nbtdRmShadingH"
    rmShadFilterV = "kernel_nbtdRmShadingV"
    
    Dim fbShadingPlane As CImgPlane
    Set fbShadingPlane = TheIDP.PlaneBank.Item("FLAT FIELD BLUE")

    Dim fbluePlane As CImgPlane
    Call GetFreePlane(fbluePlane, fbShadingPlane.planeGroup, idpDepthF32, False, "fmyBluePlane")

    Call Copy(inBluePlane, "BBAYER_FULL", EEE_COLOR_FLAT, fbluePlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)

    Call Divide(fbluePlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, fbShadingPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, fbluePlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    Call MultiplyConst(fbluePlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, 800, fbluePlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    Call Copy(fbluePlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, bluePlane, "NB_C2Z2D", EEE_COLOR_FLAT)

    Dim workPlane As CImgPlane
    Call GetFreePlane(workPlane, bluePlane.planeGroup, idpDepthS16, False, "workPlane")
    Call ExpandZoneEdgeValueH(bluePlane, "NB_C2Z2D", "Expand1", "Expand2", 7)
    Call Convolution(bluePlane, "NB_C2Z2D", EEE_COLOR_FLAT, workPlane, "NB_C2Z2D", EEE_COLOR_FLAT, "kernel_nbtdInitAntiAliasH")
    Call ExpandZoneEdgeValueV(workPlane, "NB_C2Z2D", "Expand1", "Expand2", 7)
    Call Convolution(workPlane, "NB_C2Z2D", EEE_COLOR_FLAT, bluePlane, "NB_C2Z2D", EEE_COLOR_FLAT, "kernel_nbtdInitAntiAliasV")
    
    Dim sjPlane As CImgPlane
    Call GetFreePlane(sjPlane, "c12vmcu", idpDepthS16, , "nbsjtpPlane")
    Call MultiMean(bluePlane, "NB_C2Z2D", EEE_COLOR_FLAT, sjPlane, "C_NB_DATA", EEE_COLOR_FLAT, idpMultiMeanFuncMean, COMP_PRIMARY_TO_BHLOCAL, COMP_PRIMARY_TO_BHLOCAL)

'    Dim sjMean(nSite) As Double
    Call Average(sjPlane, "C_NB_DATA", EEE_COLOR_FLAT, sjMean)
    
    Dim startCol_dummy As Long, StartRow_dummy As Long, AreaWidth As Long, AreaHeight As Long

    TheHdw.IDP.GetPMDInfo "C_NB_DATA" & "_S", startCol_dummy, StartRow_dummy, AreaWidth, AreaHeight, False

    Dim i As Long
    Dim sjMax(nSite) As Double
    Dim sjMin(nSite) As Double
    Dim sjCenter(nSite) As Double
    Call MinMax(sjPlane, "C_NB_DATA", EEE_COLOR_FLAT, sjMin, sjMax)
    For i = 0 To nSite
        If TheExec.sites.site(i).Active Then
            sjCenter(i) = (sjMax(i) + sjMin(i)) / 2
        End If
    Next i
    
    Dim acPlane As CImgPlane
    Call GetFreePlane(acPlane, sjPlane.planeGroup, idpDepthS16, False, "acPlane")
    Call SubtractConst(sjPlane, "C_NB_DATA", idpColorFlat, sjCenter, acPlane, "C_NB_DATA", idpColorFlat)

    Const IGXLBitDepth As Long = 15

'    Dim multiplyFactors(nSite) As Double
    Dim divideFactors(nSite) As Double
    Dim acAmpLog(nSite) As Double
    For i = 0 To nSite
        If TheExec.sites.site(i).Active Then
            acAmpLog(i) = Log(sjMax(i) - sjCenter(i) + 1) / Log(2)
            If acAmpLog(i) + (Log(BHShadingFilterTap) / Log(2)) * 3 > IGXLBitDepth Then
                multiplyFactors(i) = BHShadingFilterTap
                divideFactors(i) = BHShadingFilterTap
            Else
                multiplyFactors(i) = BHShadingFilterTap ^ 2
                divideFactors(i) = 1
            End If
        End If
    Next i

    Dim siteStatus As LoopStatus
    Dim curSite As Long
    Dim gainedAcPlane As CImgPlane
    Call GetFreePlane(gainedAcPlane, sjPlane.planeGroup, idpDepthS16, False, "gainedAcPlane")
    Call MultiplyConst(acPlane, "C_NB_DATA", EEE_COLOR_FLAT, multiplyFactors, gainedAcPlane, "C_NB_DATA", EEE_COLOR_FLAT)

    Call GetFreePlane(workPlane, sjPlane.planeGroup, idpDepthS16, False, "workPlane")

    Call ExpandZoneEdgeValueH(acPlane, "C_NB_DATA", "C_NB_Expand1", "C_NB_Expand2", expSize)
    Call Convolution(acPlane, "C_NB_FULL", EEE_COLOR_FLAT, workPlane, "C_NB_FULL", EEE_COLOR_FLAT, rmShadFilterH)

    Call ExpandZoneEdgeValueV(workPlane, "C_NB_DATA", "C_NB_Expand1", "C_NB_Expand2", expSize)
    Call Convolution(workPlane, "C_NB_FULL", EEE_COLOR_FLAT, acPlane, "C_NB_FULL", EEE_COLOR_FLAT, rmShadFilterV)
    Call DivideConst(acPlane, "C_NB_DATA", EEE_COLOR_FLAT, divideFactors, acPlane, "C_NB_DATA", EEE_COLOR_FLAT)
    
    Call ExpandZoneEdgeValueH(acPlane, "C_NB_DATA", "C_NB_Expand1", "C_NB_Expand2", expSize)
    Call Convolution(acPlane, "C_NB_FULL", EEE_COLOR_FLAT, workPlane, "C_NB_FULL", EEE_COLOR_FLAT, rmShadFilterH, BHShadingFilterTap)
    Call ExpandZoneEdgeValueV(workPlane, "C_NB_DATA", "C_NB_Expand1", "C_NB_Expand2", expSize)
    Call Convolution(workPlane, "C_NB_FULL", EEE_COLOR_FLAT, acPlane, "C_NB_FULL", EEE_COLOR_FLAT, rmShadFilterV, BHShadingFilterTap)
    
    Call ExpandZoneEdgeValueH(acPlane, "C_NB_DATA", "C_NB_Expand1", "C_NB_Expand2", expSize)
    Call Convolution(acPlane, "C_NB_FULL", EEE_COLOR_FLAT, workPlane, "C_NB_FULL", EEE_COLOR_FLAT, rmShadFilterH, BHShadingFilterTap)
    Call ExpandZoneEdgeValueV(workPlane, "C_NB_DATA", "C_NB_Expand1", "C_NB_Expand2", expSize)
    Call Convolution(workPlane, "C_NB_FULL", EEE_COLOR_FLAT, acPlane, "C_NB_FULL", EEE_COLOR_FLAT, rmShadFilterV, BHShadingFilterTap)

    Dim noShadPlane As CImgPlane
    Call GetFreePlane(noShadPlane, sjPlane.planeGroup, idpDepthS16, False, "noShadPlane")
    Call Subtract(gainedAcPlane, "C_NB_DATA", EEE_COLOR_FLAT, acPlane, "C_NB_DATA", EEE_COLOR_FLAT, noShadPlane, "C_NB_DATA", EEE_COLOR_FLAT)

    Call ExpandZoneEdgeValueH(noShadPlane, "C_NB_DATA", "C_NB_Expand1", "C_NB_Expand2", 7)
    Call Convolution(noShadPlane, "C_NB_FULL", EEE_COLOR_FLAT, workPlane, "C_NB_FULL", EEE_COLOR_FLAT, "kernel_nbtdFreqSelectH")
    Call ExpandZoneEdgeValueV(workPlane, "C_NB_DATA", "C_NB_Expand1", "C_NB_Expand2", 7)
    Call Convolution(workPlane, "C_NB_FULL", EEE_COLOR_FLAT, noShadPlane, "C_NB_FULL", EEE_COLOR_FLAT, "kernel_nbtdFreqSelectV")
    
    Call ExpandZoneEdgeValueH(noShadPlane, "C_NB_DATA", "C_NB_Expand1", "C_NB_Expand2", 7)
    Call Convolution(noShadPlane, "C_NB_FULL", EEE_COLOR_FLAT, workPlane, "C_NB_FULL", EEE_COLOR_FLAT, "kernel_nbtdFreqSelectH")
    Call ExpandZoneEdgeValueV(workPlane, "C_NB_DATA", "C_NB_Expand1", "C_NB_Expand2", 7)
    Call Convolution(workPlane, "C_NB_FULL", EEE_COLOR_FLAT, noShadPlane, "C_NB_FULL", EEE_COLOR_FLAT, "kernel_nbtdFreqSelectV")

'    Dim hDifPlane As CImgPlane
'    Dim vDifPlane As CImgPlane
'    Call GetFreePlane(hDifPlane, sjPlane.PlaneGroup, idpDepthS16, False, "hDifPlane")
'    Call GetFreePlane(vDifPlane, sjPlane.PlaneGroup, idpDepthS16, False, "vDifPlane")
    Call SubColumns(noShadPlane, "C_NB_DATA", EEE_COLOR_FLAT, hDifPlane, "C_NB_DATA", EEE_COLOR_FLAT, BHDifSize)
    Call SubRows(noShadPlane, "C_NB_DATA", EEE_COLOR_FLAT, vDifPlane, "C_NB_DATA", EEE_COLOR_FLAT, BHDifSize)

End Sub


Private Sub ExpandZoneEdgeValueH( _
    ByRef srcPlane As CImgPlane, _
    ByVal pZone As String, _
    ByVal modifyZone1 As String, _
    ByVal modifyZone2 As String, _
    ByVal expandBit As Long)
    
    Dim i As Long
    Const MODIFY_WIDTH As Long = 1
    
    '===== ゾーン情報取得 =============================
    Dim startCol As Long, startRow As Long, width As Long, height As Long
    Dim endCol As Long, endRow As Long
    TheHdw.IDP.GetPMDInfo pZone & "_S", startCol, startRow, width, height, False
    endCol = startCol + width - 1
    endRow = startRow + height - 1
    
    '===== 作業用プレーン確保 =====
    Dim workPlane As CImgPlane
    Call GetFreePlane(workPlane, srcPlane.planeGroup, idpDepthS16, False, "workPlane")
    
    '===== 拡張作業 =====
    '----- 左側
    modifyZone1 = modifyZone1
    modifyZone2 = modifyZone2
    TheHdw.IDP.ModifySubPMD modifyZone1 & "_S", startCol, startRow, MODIFY_WIDTH, height
    For i = 1 To expandBit
        TheHdw.IDP.ModifySubPMD modifyZone2 & "_S", startCol - i, startRow, MODIFY_WIDTH, height
        Call Copy(srcPlane, modifyZone1, EEE_COLOR_FLAT, workPlane, modifyZone2, EEE_COLOR_FLAT)
    Next i
    TheHdw.IDP.ModifySubPMD modifyZone2 & "_S", startCol - expandBit, startRow, expandBit, height
    Call Copy(workPlane, modifyZone2, EEE_COLOR_FLAT, srcPlane, modifyZone2, EEE_COLOR_FLAT)
    
    '----- 右側
    TheHdw.IDP.ModifySubPMD modifyZone1 & "_S", endCol, startRow, MODIFY_WIDTH, height
    For i = 1 To expandBit
        TheHdw.IDP.ModifySubPMD modifyZone2 & "_S", endCol + i, startRow, MODIFY_WIDTH, height
        Call Copy(srcPlane, modifyZone1, EEE_COLOR_FLAT, workPlane, modifyZone2, EEE_COLOR_FLAT)
    Next i
    TheHdw.IDP.ModifySubPMD modifyZone2 & "_S", endCol + 1, startRow, expandBit, height
    Call Copy(workPlane, modifyZone2, EEE_COLOR_FLAT, srcPlane, modifyZone2, EEE_COLOR_FLAT)
    
    Call ReleasePlane(workPlane)
    
End Sub

Private Sub ExpandZoneEdgeValueV( _
    ByRef srcPlane As CImgPlane, _
    ByVal pZone As String, _
    ByVal modifyZone1 As String, _
    ByVal modifyZone2 As String, _
    ByVal expandBit As Long)
    
    Dim i As Long
    Const MODIFY_WIDTH As Long = 1
    
    '===== ゾーン情報取得 =============================
    Dim startCol As Long, startRow As Long, width As Long, height As Long
    Dim endCol As Long, endRow As Long
    TheHdw.IDP.GetPMDInfo pZone & "_S", startCol, startRow, width, height, False
    endCol = startCol + width - 1
    endRow = startRow + height - 1
    
    '===== 作業用プレーン確保 =====
    Dim workPlane As CImgPlane
    Call GetFreePlane(workPlane, srcPlane.planeGroup, idpDepthS16, False, "workPlane")
    
    '===== 拡張作業 =====
    '----- 上側
    TheHdw.IDP.ModifySubPMD modifyZone1 & "_S", startCol, startRow, width, MODIFY_WIDTH
    For i = 1 To expandBit
        TheHdw.IDP.ModifySubPMD modifyZone2 & "_S", startCol, startRow - i, width, MODIFY_WIDTH
        Call Copy(srcPlane, modifyZone1, EEE_COLOR_FLAT, workPlane, modifyZone2, EEE_COLOR_FLAT)
    Next i
    TheHdw.IDP.ModifySubPMD modifyZone2 & "_S", startCol, startRow - expandBit, width, expandBit
    Call Copy(workPlane, modifyZone2, EEE_COLOR_FLAT, srcPlane, modifyZone2, EEE_COLOR_FLAT)
    
    '----- 下側
    TheHdw.IDP.ModifySubPMD modifyZone1 & "_S", startCol, endRow, width, MODIFY_WIDTH
    For i = 1 To expandBit
        TheHdw.IDP.ModifySubPMD modifyZone2 & "_S", startCol, endRow + i, width, MODIFY_WIDTH
        Call Copy(srcPlane, modifyZone1, EEE_COLOR_FLAT, workPlane, modifyZone2, EEE_COLOR_FLAT)
    Next i
    TheHdw.IDP.ModifySubPMD modifyZone2 & "_S", startCol, endRow + 1, width, expandBit
    Call Copy(workPlane, modifyZone2, EEE_COLOR_FLAT, srcPlane, modifyZone2, EEE_COLOR_FLAT)
    
    Call ReleasePlane(workPlane)
    
End Sub

Private Sub B_N_Hakimura( _
    ByRef hDifPlane As CImgPlane, ByRef vDifPlane As CImgPlane, ByVal pSlice As Double, ByRef sjMean() As Double, ByRef multiplyFactors() As Double, returnResult_mf() As Double, returnResult_af() As Double)

    Dim HiLimit(nSite) As Double, LoLimit(nSite) As Double
    Dim i As Long

    For i = 0 To nSite
        If TheExec.sites.site(i).Active = True Then
            HiLimit(i) = Int(sjMean(i) * multiplyFactors(i) * pSlice) - 1
            LoLimit(i) = HiLimit(i) * (-1)
        End If
    Next i

    Dim mnbsjtp2(nSite) As Double, mnbsjtp2H(nSite) As Double, mnbsjtp2V(nSite) As Double
    Call Count(hDifPlane, "C_NB_HJUDGE_CZ", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mnbsjtp2H, "FLG_NBSJTP2_H")
    Call Count(vDifPlane, "C_NB_VJUDGE_CZ", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mnbsjtp2V, "FLG_NBSJTP2_V")

    Dim FlgPlane As CImgPlane
    Call GetFreePlane(FlgPlane, hDifPlane.planeGroup, idpDepthS16, True, "flgPlane")
    Call SharedFlagOr(vDifPlane.planeGroup, "C_NB_DATA", "FLG_NBSJTP2_HV", "FLG_NBSJTP2_H", "FLG_NBSJTP2_V")
    Call FlagCopy(FlgPlane, "C_NB_DATA", "FLG_NBSJTP2_HV")

    Call Count(FlgPlane, "C_NB_DATA", EEE_COLOR_FLAT, idpCountAbove, 0, 0, idpLimitExclude, mnbsjtp2, "FLG_NBSJTP2")

    Dim hDifMin(nSite) As Double
    Dim hDifMax(nSite) As Double
    Dim hDifResult(nSite) As Double
    Call MinMax(hDifPlane, "C_NB_HJUDGE_CZ", EEE_COLOR_FLAT, hDifMin, hDifMax)
    For i = 0 To nSite
        If TheExec.sites.site(i).Active Then
            If Abs(hDifMin(i)) > Abs(hDifMax(i)) Then
                hDifResult(i) = Abs(hDifMin(i))
            Else
                hDifResult(i) = Abs(hDifMax(i))
            End If
        End If
    Next i
    
    Dim vDifMin(nSite) As Double
    Dim vDifMax(nSite) As Double
    Dim vDifResult(nSite) As Double
    Call MinMax(vDifPlane, "C_NB_VJUDGE_CZ", EEE_COLOR_FLAT, vDifMin, vDifMax)
    For i = 0 To nSite
        If TheExec.sites.site(i).Active Then
            If Abs(vDifMin(i)) > Abs(vDifMax(i)) Then
                vDifResult(i) = Abs(vDifMin(i))
            Else
                vDifResult(i) = Abs(vDifMax(i))
            End If
        End If
    Next i

    Dim startCol_dummy As Long, StartRow_dummy As Long, AreaWidth As Long, AreaHeight As Long

    TheHdw.IDP.GetPMDInfo "C_NB_DATA" & "_S", startCol_dummy, StartRow_dummy, AreaWidth, AreaHeight, False

    Dim difResult(nSite) As Double
    Dim sjGainedMean(nSite) As Double

    For i = 0 To nSite
        If TheExec.sites.site(i).Active Then
            If hDifResult(i) > vDifResult(i) Then
                difResult(i) = hDifResult(i)
            Else
                difResult(i) = vDifResult(i)
            End If
            returnResult_mf(i) = mf_div(difResult(i), sjMean(i) * multiplyFactors(i), 1000)
            returnResult_af(i) = mnbsjtp2(i) / (AreaWidth * AreaHeight)
        End If
    Next i

End Sub

Private Sub B_G_Hakimura( _
    ByRef hDifPlane As CImgPlane, ByRef vDifPlane As CImgPlane, ByVal pSlice As Double, ByRef sjMean() As Double, ByRef multiplyFactors() As Double, returnResult_mf() As Double, returnResult_af() As Double)

    Dim HiLimit(nSite) As Double, LoLimit(nSite) As Double
    Dim i As Long

    For i = 0 To nSite
        If TheExec.sites.site(i).Active = True Then
            HiLimit(i) = Int(sjMean(i) * multiplyFactors(i) * pSlice) - 1
            LoLimit(i) = HiLimit(i) * (-1)
        End If
    Next i

    Call WritePixel(hDifPlane, "C_NB_HCLEAR_CZ", EEE_COLOR_FLAT, 0)
    Call WritePixel(vDifPlane, "C_NB_VCLEAR_CZ", EEE_COLOR_FLAT, 0)

    Dim mnbsjtpf2d(nSite) As Double, mnbsjtpf2dH(nSite) As Double, mnbsjtpf2dV(nSite) As Double
    Call Count(hDifPlane, "C_NB_HJUDGE_FZ", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mnbsjtpf2dH, "FLG_NBSJTPF2D_H")
    Call Count(vDifPlane, "C_NB_VJUDGE_FZ", EEE_COLOR_FLAT, idpCountOutside, LoLimit, HiLimit, idpLimitExclude, mnbsjtpf2dV, "FLG_NBSJTPF2D_V")

    Dim FlgPlane As CImgPlane
    Call GetFreePlane(FlgPlane, hDifPlane.planeGroup, idpDepthS16, True, "flgPlane")
    Call SharedFlagOr(vDifPlane.planeGroup, "C_NB_DATA", "FLG_NBSJTPF2D_HV", "FLG_NBSJTPF2D_H", "FLG_NBSJTPF2D_V")
    Call FlagCopy(FlgPlane, "C_NB_DATA", "FLG_NBSJTPF2D_HV")

    Call Count(FlgPlane, "C_NB_DATA", EEE_COLOR_FLAT, idpCountAbove, 0, 0, idpLimitExclude, mnbsjtpf2d, "FLG_NBSJTPF2D")

    Dim hDifMin(nSite) As Double
    Dim hDifMax(nSite) As Double
    Dim hDifResult(nSite) As Double
    
    Call MinMax(hDifPlane, "C_NB_HJUDGE_FZ", EEE_COLOR_FLAT, hDifMin, hDifMax)
    For i = 0 To nSite
        If TheExec.sites.site(i).Active Then
            If Abs(hDifMin(i)) > Abs(hDifMax(i)) Then
                hDifResult(i) = Abs(hDifMin(i))
            Else
                hDifResult(i) = Abs(hDifMax(i))
            End If
        End If
    Next i

    Dim vDifMin(nSite) As Double
    Dim vDifMax(nSite) As Double
    Dim vDifResult(nSite) As Double
    
    Call MinMax(vDifPlane, "C_NB_VJUDGE_FZ", EEE_COLOR_FLAT, vDifMin, vDifMax)
    For i = 0 To nSite
        If TheExec.sites.site(i).Active Then
            If Abs(vDifMin(i)) > Abs(vDifMax(i)) Then
                vDifResult(i) = Abs(vDifMin(i))
            Else
                vDifResult(i) = Abs(vDifMax(i))
            End If
        End If
    Next i

    Dim startCol_dummy As Long, StartRow_dummy As Long, AreaWidth As Long, AreaHeight As Long

    TheHdw.IDP.GetPMDInfo "C_NB_DATA" & "_S", startCol_dummy, StartRow_dummy, AreaWidth, AreaHeight, False
    
    Dim difResult(nSite) As Double
    Dim sjGainedMean(nSite) As Double

    For i = 0 To nSite
        If TheExec.sites.site(i).Active Then
            If hDifResult(i) > vDifResult(i) Then
                difResult(i) = hDifResult(i)
            Else
                difResult(i) = vDifResult(i)
            End If
            returnResult_mf(i) = mf_div(difResult(i), sjMean(i) * multiplyFactors(i), 1000)
            returnResult_af(i) = mnbsjtpf2d(i) / (AreaWidth * AreaHeight)
        End If
    Next i

End Sub

Public Sub StdBHCM_GetMuraParameters()
    Dim MuraCol As Collection
    Set MuraCol = New Collection

    Call stdBHCM_GetMuraParam(MuraCol)
    
    Call StdBHCM_SetModParam(MuraCol)
    
    Set MuraCol = Nothing
End Sub

Public Function StdBHCM_SetModParam(ByRef MuraCol As Collection)
    With MuraCol
        COMP_RGB_TO_Y = .Item("ConstantsCompressionSize_RGBtoYLINE")
        BHDifSize = .Item("ConstantsB_HAKIDifferentialDistance")
        BHShadingFilterTap = .Item("ConstantsB_HAKIShadingFilterTap")
        COMP_PRIMARY_TO_BHLOCAL = .Item("ConstantsB_HAKICompressPixel")
    End With
    
    With MuraCol
        Nbtd_FlatFieldPath = .Item("PATHFlatFieldPath")
    End With
   
End Function


Private Sub stdBHCM_GetMuraParam(ByRef MuraCol As Collection)
    
    Dim Loopi As Long
    Dim LoopB As Long
    Dim buf As String
    Dim bufstr As String
    Dim StartLow As Long
    
    StartLow = 5
    
    Loopi = StartLow
    'C列　"Parameter Name" 検索
    Do Until Worksheets("BHAKI_Mura Parameters").Range("C" & Loopi) = ""
        LoopB = Loopi
        
        'B列 "Mura Item" 検索
        Do While Worksheets("BHAKI_Mura Parameters").Range("B" & LoopB) = ""
            LoopB = LoopB - 1
            
            If LoopB < StartLow Then
                Exit Do
            End If
        Loop
        'D列 "Value"取得 Keyは"Mura Item"&"Parameter Name"となる。
        MuraCol.Add Item:=Worksheets("BHAKI_Mura Parameters").Range("D" & Loopi), key:=Worksheets("BHAKI_Mura Parameters").Range("B" & LoopB) & Worksheets("BHAKI_Mura Parameters").Range("C" & Loopi)

        Loopi = Loopi + 1
        If Loopi > 1000 Then
            Exit Do
        End If
    Loop

End Sub

Public Sub NBSJ_preProcess(InputPlane As CImgPlane, InputZone As Variant, ByRef sjMean() As Double)

    '========== BAYER SEPARATION ==================================
    Dim bayerBluePlane As CImgPlane
    Call GetFreePlane(bayerBluePlane, "bbayer", idpDepthS16, False, "bluePlane for bayer")
    
    Dim tmpBayerPlane As CImgPlane
    Call GetFreePlane(tmpBayerPlane, "allbayer", idpDepthS16, False, "Clamp Image (Bayer plane)")
    Call Copy(InputPlane, InputZone, EEE_COLOR_FLAT, tmpBayerPlane, "ALLBAYER_ZONE2D", EEE_COLOR_FLAT)
    
    Call MultiMean(tmpBayerPlane, "ALLBAYER_ZONE2D", "B", bayerBluePlane, "BBAYER_FULL", "B", idpMultiMeanFuncMean, StdCM_CompRGB2Y, StdCM_CompRGB2Y)

    Dim multiplyFactors(nSite) As Double
    
    Dim hDifPlane As CImgPlane
    Dim vDifPlane As CImgPlane
    Call GetFreePlane(hDifPlane, "c12vmcu", idpDepthS16, False, "hDifPlane")
    Call GetFreePlane(vDifPlane, "c12vmcu", idpDepthS16, False, "vDifPlane")
    
    Dim bluePlane As CImgPlane
    Call GetFreePlane(bluePlane, InputPlane.planeGroup, idpDepthS16, False, "myBluePlane")
    
    Call Pre_BHakimura(bayerBluePlane, bluePlane, sjMean, multiplyFactors, hDifPlane, vDifPlane)

    Call StdCM_RegisteredValue("BHAKI_MULTIPLYFACTORS", multiplyFactors)
    Call TheIDP.PlaneBank.Add("BHAKI_PRE_HDIFF_PLANE", hDifPlane)
    Call TheIDP.PlaneBank.Add("BHAKI_PRE_VDIFF_PLANE", vDifPlane)

End Sub

Public Sub NBSJ_inner(Slice As Variant, ByRef sjMean() As Double, ByRef result() As Double)

    Dim site As Long

    Dim ret_tmp1(nSite) As Double
    Dim ret_tmp2(nSite) As Double

    Dim multiplyFactors() As Double
    TheResult.GetResult "BHAKI_MULTIPLYFACTORS", multiplyFactors

    Dim hDifPlane As CImgPlane, vDifPlane As CImgPlane
    
    Call GetRegisteredPlane("BHAKI_PRE_HDIFF_PLANE", hDifPlane, "BH_hdifPlane")
    Call GetRegisteredPlane("BHAKI_PRE_VDIFF_PLANE", vDifPlane, "BH_vdifPlane")
    
    Dim workPlane_h As CImgPlane
    Dim workPlane_v As CImgPlane
    Call GetFreePlane(workPlane_h, "c12vmcu", idpDepthS16, False, "WorkPlane_h")
    Call GetFreePlane(workPlane_v, "c12vmcu", idpDepthS16, False, "WorkPlane_v")
    Call Copy(hDifPlane, "C_NB_FULL", EEE_COLOR_ALL, workPlane_h, "C_NB_FULL", EEE_COLOR_ALL)
    Call Copy(vDifPlane, "C_NB_FULL", EEE_COLOR_ALL, workPlane_v, "C_NB_FULL", EEE_COLOR_ALL)

    Call B_N_Hakimura(workPlane_h, workPlane_v, Slice, sjMean, multiplyFactors, ret_tmp1, ret_tmp2)

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            result(CMReturnType.rMAX, site) = ret_tmp1(site)
            result(CMReturnType.rNUM, site) = ret_tmp2(site)
        End If
    Next site

End Sub

Public Sub NBSJ_outer(Slice As Variant, ByRef sjMean() As Double, ByRef result() As Double)

    Dim site As Long

    Dim ret_tmp1(nSite) As Double
    Dim ret_tmp2(nSite) As Double

    Dim multiplyFactors() As Double
    TheResult.GetResult "BHAKI_MULTIPLYFACTORS", multiplyFactors

    Dim hDifPlane As CImgPlane, vDifPlane As CImgPlane
    
    Call GetRegisteredPlane("BHAKI_PRE_HDIFF_PLANE", hDifPlane, "BH_hdifPlane")
    Call GetRegisteredPlane("BHAKI_PRE_VDIFF_PLANE", vDifPlane, "BH_vdifPlane")
    
    Dim workPlane_h As CImgPlane
    Dim workPlane_v As CImgPlane
    Call GetFreePlane(workPlane_h, "c12vmcu", idpDepthS16, False, "WorkPlane_h")
    Call GetFreePlane(workPlane_v, "c12vmcu", idpDepthS16, False, "WorkPlane_v")
    Call Copy(hDifPlane, "C_NB_FULL", EEE_COLOR_ALL, workPlane_h, "C_NB_FULL", EEE_COLOR_ALL)
    Call Copy(vDifPlane, "C_NB_FULL", EEE_COLOR_ALL, workPlane_v, "C_NB_FULL", EEE_COLOR_ALL)

    Call B_G_Hakimura(workPlane_h, workPlane_v, Slice, sjMean, multiplyFactors, ret_tmp1, ret_tmp2)

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            result(CMReturnType.rMAX, site) = ret_tmp1(site)
            result(CMReturnType.rNUM, site) = ret_tmp2(site)
        End If
    Next site

End Sub

'-------------------------------------------------------------------------------------
'# Name         CM_RGBandIRseparate
'# Purpose
'#              IR画素配列でのクラシックムラの前処理。
'#              [Preprocessing function for bayer color map device with one of the "G" colors
'#              replaced with "IR"(infrared or some color that cannot be used for Mura evaluation).]
'#
'# Restrictions 通常は、以下のカラーマップを想定している。
'#                   R   Gr
'#                   IR  B
'#              但し、"IR"はムラ処理に寄与しないカラーとして処理する(無視する色)ため、
'#              IR(赤外)でなくても構いません。
'#              [This function is restricted to the following color map.
'#                   R   Gr
'#                   IR  B
'#              Note that "IR(Infrared)" means a color ignored in this function. Therefore,
'#              this can be a color other than infrared (e.g.: Black, etc).]
'#              　より一般的には、カラーマップ配列中が以下の条件をみたすこと
'#                      1). "R"、"Gr"、"B"がそれぞれ1つ以上存在する
'#                      2). マップ中の"R", "Gr", "B"の個数は同数
'#                      3). マップ中の"R"のみを抜き出し左詰めに並べると、正方形になる
'#              　[More generally, the following items on color map must be satisfied.
'#                      a.  One or more "R" must exist. Same for "Gr" and "B".
'#                      b.  The number of "R", "Gr", and "B" must be the same.
'#                      c.  When left aligned, "R" cells must form quadrate.
'# History      First drafted by T.Morimoto 2013-12-09 for IMX288.
'-------------------------------------------------------------------------------------
Public Sub CM_RGBandIRseparate(InputPlane As CImgPlane, InputZone As Variant, ByRef result() As Double)

    Dim site As Long

    '========== BAYER SEPARATION ==================================
    Dim bayerRedPlane As CImgPlane
    Dim bayerGreenPlane As CImgPlane
    Dim bayerBluePlane As CImgPlane
    Call GetFreePlane(bayerRedPlane, "rbayer", idpDepthS16, False, "redPlane for bayer")
    Call GetFreePlane(bayerGreenPlane, "gbayer", idpDepthS16, False, "greenPlane for bayer")
    Call GetFreePlane(bayerBluePlane, "bbayer", idpDepthS16, False, "bluePlane for bayer")
    
    Dim tmpBayerPlane As CImgPlane
    Call GetFreePlane(tmpBayerPlane, "allbayer", idpDepthS16, False, "Clamp Image (Bayer plane)")
    Call Copy(InputPlane, InputZone, EEE_COLOR_FLAT, tmpBayerPlane, "ALLBAYER_ZONE2D", EEE_COLOR_FLAT)
    
    '通常のベイヤー処理とはここだけが違います。あとは全部一緒。
    'Only difference from the normal bayer process is the function called here.
    Call StdCM_SeparateRGBandIRforBayer(tmpBayerPlane, "ALLBAYER_ZONE2D", _
                                        bayerRedPlane, "RBAYER_FULL", _
                                        bayerGreenPlane, "GBAYER_FULL", _
                                        bayerBluePlane, "BBAYER_FULL")

   
    Dim rRawPlane As CImgPlane
    Dim gRawPlane As CImgPlane
    Dim bRawPlane As CImgPlane
    Call GetFreePlane(rRawPlane, "pyline", idpDepthF32, False, "Red Raw Image")
    Call GetFreePlane(gRawPlane, "pyline", idpDepthF32, False, "Green Raw Image")
    Call GetFreePlane(bRawPlane, "pyline", idpDepthF32, False, "Blue Raw Image")
    Call Copy(bayerRedPlane, "RBAYER_FULL", EEE_COLOR_FLAT, rRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    Call Copy(bayerGreenPlane, "GBAYER_FULL", EEE_COLOR_FLAT, gRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    Call Copy(bayerBluePlane, "BBAYER_FULL", EEE_COLOR_FLAT, bRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)

    Call ReleasePlane(bayerRedPlane)
    Call ReleasePlane(bayerGreenPlane)
    Call ReleasePlane(bayerBluePlane)
    
    Dim redPlane As CImgPlane
    Dim greenPlane As CImgPlane
    Dim bluePlane As CImgPlane
    Call GetFreePlane(redPlane, "pyline", idpDepthF32, False, "red plane")
    Call GetFreePlane(greenPlane, "pyline", idpDepthF32, False, "green plane")
    Call GetFreePlane(bluePlane, "pyline", idpDepthF32, False, "blue plane")
    If IsClsMuraFlatFieldingOn Then
        Dim frShadingPlane As CImgPlane
        Dim fgShadingPlane As CImgPlane
        Dim fbShadingPlane As CImgPlane
        Set frShadingPlane = TheIDP.PlaneBank.Item("FLAT FIELD RED")
        Set fgShadingPlane = TheIDP.PlaneBank.Item("FLAT FIELD GREEN")
        Set fbShadingPlane = TheIDP.PlaneBank.Item("FLAT FIELD BLUE")
    
        Call Divide(rRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, _
                    frShadingPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, _
                    redPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
        Call Divide(gRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, _
                    fgShadingPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, _
                    greenPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
        Call Divide(bRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, _
                    fbShadingPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, _
                    bluePlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)

        Call ReleasePlane(frShadingPlane)
        Call ReleasePlane(fgShadingPlane)
        Call ReleasePlane(fbShadingPlane)
    Else
        Call Copy(rRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, redPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
        Call Copy(gRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, greenPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
        Call Copy(bRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, bluePlane, "YLINE_ZONE2D", EEE_COLOR_FLAT)
    End If

    Call ReleasePlane(rRawPlane)
    Call ReleasePlane(gRawPlane)
    Call ReleasePlane(bRawPlane)
    
    '以降の処理で必要なので、Ｒ，Ｇ，Ｂの画像をプレーンバンクに登録
    Call TheIDP.PlaneBank.Add("MURA_RED_SEP_PLANE", redPlane, , True)
    Call TheIDP.PlaneBank.Add("MURA_GREEN_SEP_PLANE", greenPlane, , True)
    Call TheIDP.PlaneBank.Add("MURA_BLUE_SEP_PLANE", bluePlane, , True)
    
    '========== MEAN VALUES OF R/G/B and Y ==================================
    Dim redMean(nSite) As Double
    Dim greenMean(nSite) As Double
    Dim blueMean(nSite) As Double
    Call Average(redPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, redMean)
    Call Average(greenPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, greenMean)
    Call Average(bluePlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, blueMean)
    
    For site = 0 To nSite
        result(rgbColorArray.red, site) = redMean(site)
        result(rgbColorArray.green, site) = greenMean(site)
        result(rgbColorArray.blue, site) = blueMean(site)
    Next site

End Sub

'-------------------------------------------------------------------------------------
'# Name         StdCM_SeparateRGBandIRforBayer
'# Purpose
'#              IR画素配列(ベイヤー配列のGbがIR画素もしくは、ムラ評価では使用できないカラーとなった
'#              ケース)でのクラシックムラのRGB分割処理関数
'#              [RGB color separation function for bayer color map device with "Gb" color replaced
'#              with IR(infrared or some color that cannot be used for Mura evaluation).]
'# Restrictions Must be called from the function "CM_RGBandIRseparate".
'# History      First drafted by T.Morimoto 2013-12-09 for IMX288.
'-------------------------------------------------------------------------------------
Public Sub StdCM_SeparateRGBandIRforBayer( _
    ByRef srcPlane As CImgPlane, _
    ByVal srcZone As String, _
    ByRef redPlane As CImgPlane, _
    ByVal redZone As String, _
    ByRef greenPlane As CImgPlane, _
    ByVal greenZone As String, _
    ByRef bluePlane As CImgPlane, _
    ByVal blueZone As String)

    Call MultiMean(srcPlane, srcZone, "R", redPlane, redZone, "R", idpMultiMeanFuncMean, StdCM_CompRGB2Y, StdCM_CompRGB2Y)
    Call MultiMean(srcPlane, srcZone, "B", bluePlane, blueZone, "B", idpMultiMeanFuncMean, StdCM_CompRGB2Y, StdCM_CompRGB2Y)
    Call MultiMean(srcPlane, srcZone, "GR", greenPlane, greenZone, "GR", idpMultiMeanFuncMean, StdCM_CompRGB2Y, StdCM_CompRGB2Y)

End Sub

'-------------------------------------------------------------------------------------
'# Name         labProc_RGBandIRSeparationFF
'# Purpose
'#              IR画素配列でのL*a*b*ムラのFlat Field前処理。
'#              [Flat field preprocessing function for bayer color map device with one of the "G" colors
'#              replaced with "IR"(infrared or some color that cannot be used for Mura evaluation).]
'#
'# Restrictions 通常は、以下のカラーマップを想定している。
'#                   R   Gr
'#                   IR  B
'#              但し、"IR"はムラ処理に寄与しないカラーとして処理する(無視する色)ため、
'#              IR(赤外)でなくても構いません。
'#              [This function is restricted to the following color map.
'#                   R   Gr
'#                   IR  B
'#              Note that "IR(Infrared)" means a color ignored in this function. Therefore,
'#              this can be a color other than infrared (e.g.: Black, etc).]
'#              　より一般的には、カラーマップ配列中が以下の条件をみたすこと
'#                      1). "R"、"Gr"、"B"がそれぞれ1つ以上存在する
'#                      2). マップ中の"R", "Gr", "B"の個数は同数
'#                      3). マップ中の"R"のみを抜き出し左詰めに並べると、正方形になる
'#              　[More generally, the following items on color map must be satisfied.
'#                      a.  One or more "R" must exist. Same for "Gr" and "B".
'#                      b.  The number of "R", "Gr", and "B" must be the same.
'#                      c.  When left aligned, "R" cells must form quadrate.
'# History      First drafted by T.Morimoto 2013-12-09 for IMX288.
'-------------------------------------------------------------------------------------
Private Sub labProc_RGBandIRSeparationFF( _
    ByRef flatPlane As CImgPlane, ByVal ZONE_ZONE3 As Variant, ByVal clampZone As Variant)
    
    'Image planes
    Const MYPLANE_LOCAL As String = "pLabL"
    Const MYPLANE_BAYER As String = "pBayer"
    Const MYPLANE_BAYER_R As String = "pBayerR"
    Const MYPLANE_BAYER_G As String = "pBayerG"
    Const MYPLANE_BAYER_B As String = "pBayerB"
    Const MYPLANE_COMP_BAYER_R As String = MYPLANE_LOCAL
    Const MYPLANE_COMP_BAYER_G As String = "pcBayerG"
    Const MYPLANE_COMP_BAYER_B As String = "pcBayerB"
    
    Dim compFactor As Long
    compFactor = labProc_ReturnCompFactor
    
    'Local variables
    Dim Bclamp(nSite) As Double
    
    'OPB Clamp
    Dim clampPlane As CImgPlane
    Call GetFreePlane(clampPlane, flatPlane.planeGroup, idpDepthS16, False, "MURA/Clamp Image")
    Call Average(flatPlane, clampZone, EEE_COLOR_FLAT, Bclamp)
    Call SubtractConst(flatPlane, ZONE_ZONE3, EEE_COLOR_FLAT, Bclamp, clampPlane, ZONE_ZONE3, EEE_COLOR_FLAT)
    
    'RGB separation
    Dim bayerPlane As CImgPlane
    Call GetFreePlane(bayerPlane, MYPLANE_BAYER, idpDepthS16, False, "Bayer Plane")
    Call Copy(clampPlane, ZONE_ZONE3, EEE_COLOR_FLAT, bayerPlane, "BAYER_FULL", EEE_COLOR_FLAT)
    Call ReleasePlane(clampPlane)
    
    Dim bayerPlaneR As CImgPlane
    Call GetFreePlane(bayerPlaneR, MYPLANE_BAYER_R, idpDepthS16, False, "R extracted plane")
    Call Copy(bayerPlane, "BAYER_FULL", "R", bayerPlaneR, "BAYER_R_FULL", "R")
    
    Dim bayerPlaneG As CImgPlane
    Call GetFreePlane(bayerPlaneG, MYPLANE_BAYER_G, idpDepthS16, False, "G extracted plane")
    Call Copy(bayerPlane, "BAYER_FULL", "GR", bayerPlaneG, "BAYER_G_FULL", "GR")
    
    Dim bayerPlaneB As CImgPlane
    Call GetFreePlane(bayerPlaneB, MYPLANE_BAYER_B, idpDepthS16, False, "B extracted plane")
    Call Copy(bayerPlane, "BAYER_FULL", "B", bayerPlaneB, "BAYER_B_FULL", "B")
    Call ReleasePlane(bayerPlane)
    
    'Random noise reduction (median filter)
    Dim bayerPlaneWorkR As CImgPlane
    Call GetFreePlane(bayerPlaneWorkR, bayerPlaneR.planeGroup, idpDepthS16, False, "Work for R-extracted")
    Call Median(bayerPlaneR, "BAYER_R_FULL", EEE_COLOR_FLAT, bayerPlaneWorkR, "BAYER_R_FULL", EEE_COLOR_FLAT, 5, 1)
    Call Median(bayerPlaneWorkR, "BAYER_R_FULL", EEE_COLOR_FLAT, bayerPlaneR, "BAYER_R_FULL", EEE_COLOR_FLAT, 1, 5)
    Call ReleasePlane(bayerPlaneWorkR)
    
    Dim bayerPlaneWorkG As CImgPlane
    Call GetFreePlane(bayerPlaneWorkG, bayerPlaneG.planeGroup, idpDepthS16, False, "Work for G-extracted")
    Call Median(bayerPlaneG, "BAYER_G_FULL", EEE_COLOR_FLAT, bayerPlaneWorkG, "BAYER_G_FULL", EEE_COLOR_FLAT, 5, 1)
    Call Median(bayerPlaneWorkG, "BAYER_G_FULL", EEE_COLOR_FLAT, bayerPlaneG, "BAYER_G_FULL", EEE_COLOR_FLAT, 1, 5)
    Call ReleasePlane(bayerPlaneWorkG)
    
    Dim bayerPlaneWorkB As CImgPlane
    Call GetFreePlane(bayerPlaneWorkB, bayerPlaneB.planeGroup, idpDepthS16, False, "Work for B-extracted")
    Call Median(bayerPlaneB, "BAYER_B_FULL", EEE_COLOR_FLAT, bayerPlaneWorkB, "BAYER_B_FULL", EEE_COLOR_FLAT, 5, 1)
    Call Median(bayerPlaneWorkB, "BAYER_B_FULL", EEE_COLOR_FLAT, bayerPlaneB, "BAYER_B_FULL", EEE_COLOR_FLAT, 1, 5)
    Call ReleasePlane(bayerPlaneWorkB)

    'Multi-Mean for local
    Dim rISrcPlane As CImgPlane
    Dim gISrcPlane As CImgPlane
    Dim bISrcPlane As CImgPlane
    Dim gCompPlane As CImgPlane
    Dim bCompPlane As CImgPlane
    Call GetFreePlane(rISrcPlane, MYPLANE_LOCAL, idpDepthS16, False, "R (int) Source Plane")
    Call GetFreePlane(gISrcPlane, MYPLANE_LOCAL, idpDepthS16, False, "G (int) Source Plane")
    Call GetFreePlane(bISrcPlane, MYPLANE_LOCAL, idpDepthS16, False, "B (int) Source Plane")
    Call GetFreePlane(gCompPlane, MYPLANE_COMP_BAYER_G, idpDepthS16, False, "G Compressed Plane")
    Call GetFreePlane(bCompPlane, MYPLANE_COMP_BAYER_B, idpDepthS16, False, "B Compressed Plane")
    Call MultiMean(bayerPlaneR, "BAYER_R_ZONE2D", EEE_COLOR_FLAT, _
                   rISrcPlane, "LABZ2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, compFactor, compFactor)
    Call MultiMean(bayerPlaneG, "BAYER_G_ZONE2D", EEE_COLOR_FLAT, _
                   gCompPlane, "PCBAYER_G_Z2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, compFactor, compFactor)
    Call Copy(gCompPlane, "PCBAYER_G_Z2D", EEE_COLOR_FLAT, gISrcPlane, "LABZ2D", EEE_COLOR_FLAT)
    Call MultiMean(bayerPlaneB, "BAYER_B_ZONE2D", EEE_COLOR_FLAT, _
                   bCompPlane, "PCBAYER_B_Z2D", EEE_COLOR_FLAT, idpMultiMeanFuncMean, compFactor, compFactor)
    Call Copy(bCompPlane, "PCBAYER_B_Z2D", EEE_COLOR_FLAT, bISrcPlane, "LABZ2D", EEE_COLOR_FLAT)
    Call ReleasePlane(bayerPlaneR)
    Call ReleasePlane(bayerPlaneG)
    Call ReleasePlane(bayerPlaneB)
    Call ReleasePlane(gCompPlane)
    Call ReleasePlane(bCompPlane)
    
'直前のGとBのコピーは直接浮動小数にできないか見ること。
    'Flat Fielding
    Dim rFSrcPlane As CImgPlane
    Dim gFSrcPlane As CImgPlane
    Dim bFSrcPlane As CImgPlane
    Call GetFreePlane(rFSrcPlane, rISrcPlane.planeGroup, idpDepthF32, False, "R (Float) Source Plane")
    Call GetFreePlane(gFSrcPlane, gISrcPlane.planeGroup, idpDepthF32, False, "G (Float) Source Plane")
    Call GetFreePlane(bFSrcPlane, bISrcPlane.planeGroup, idpDepthF32, False, "B (Float) Source Plane")
    Call Copy(rISrcPlane, "LABZ2D", EEE_COLOR_FLAT, rFSrcPlane, "LABZ2D", EEE_COLOR_FLAT)
    Call Copy(gISrcPlane, "LABZ2D", EEE_COLOR_FLAT, gFSrcPlane, "LABZ2D", EEE_COLOR_FLAT)
    Call Copy(bISrcPlane, "LABZ2D", EEE_COLOR_FLAT, bFSrcPlane, "LABZ2D", EEE_COLOR_FLAT)
    Call ReleasePlane(rISrcPlane)
    Call ReleasePlane(gISrcPlane)
    Call ReleasePlane(bISrcPlane)
    
    With TheIDP.PlaneBank
        If .isExisting(PLANEBANK_FLATFIELD_R) Then Call .Delete(PLANEBANK_FLATFIELD_R)
        Call .Add(PLANEBANK_FLATFIELD_R, rFSrcPlane, True, True)
        
        If .isExisting(PLANEBANK_FLATFIELD_G) Then Call .Delete(PLANEBANK_FLATFIELD_G)
        Call .Add(PLANEBANK_FLATFIELD_G, gFSrcPlane, True, True)
        
        If .isExisting(PLANEBANK_FLATFIELD_B) Then Call .Delete(PLANEBANK_FLATFIELD_B)
        Call .Add(PLANEBANK_FLATFIELD_B, bFSrcPlane, True, True)
    End With
    
End Sub


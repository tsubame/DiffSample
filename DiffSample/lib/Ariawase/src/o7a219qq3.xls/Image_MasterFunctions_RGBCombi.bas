Attribute VB_Name = "Image_MasterFunctions_RGBCombi"
Option Explicit
        
'RGB合成によるクラシックムラ=============================================
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
Private rgbCombiRefFlatFieldMeanR(nSite) As Double
Private rgbCombiRefFlatFieldMeanG(nSite) As Double
Private rgbCombiRefFlatFieldMeanB(nSite) As Double


'Labムラ==================================================
Const PLANEBANK_FLATFIELD_R As String = "Flat Field R"
Const PLANEBANK_FLATFIELD_G As String = "Flat Field G"
Const PLANEBANK_FLATFIELD_B As String = "Flat Field B"

'---- Path to FF
Private LabFlatFieldPath As String

'---- Kernel Taps
Private LabKernel_LowPassH As String
Private LabKernel_LowPassV As String

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

Public Enum RGBCombiCMReturnType
  rNUM = 1
  rMAX = 2 '配列の要素数としても使用しているので、MAXを最大とすること
End Enum

Public Enum RGBCombiRGBColorArray
  red = 0
  green = 1
  blue = 2
End Enum

'-------------------------------------------------------------------------------------
'# Name         RGBCombiCM_Initialize
'# Purpose
'#
'#
'# Restrictions
'# History      First drafted by T.Morimoto 2013-12-09 for IMX288.
'-------------------------------------------------------------------------------------
Public Sub RGBCombiCM_Initialize(ByVal pType As String, ByVal ZONE_FULL As Variant, ByVal ZONE_ZONE3 As Variant, ByVal clampZone As Variant, ByVal paramSheetName As String)

    'Local variables
    Dim site As Long                'For site loop

    Dim Flg_Active(nSite) As Long

    If TheIDP.KernelManager.IsExist("kernel_RGBCombi3x3") = True Then Exit Sub  'CHECK

     'ALL SITE ACTIVE
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Flg_Active(site) = 1
        End If
    Next site
    Call TheExec.sites.SetAllActive(True)
    
    If Not TheIDP.LUTManager.IsExist("lut_arctan") Then Call lut_v30

    'To read parameters
    Call RGBCombiCM_GetMuraParameters(paramSheetName)
     'To define kernels.
    Call ker_RGBCombiCM_v30
    'Flat Field Mode
    If IsClsMuraFlatFieldingOn Then
        
        Call RGBCombiCM_MakeFlatFieldImagePlane(pType, ZONE_FULL, ZONE_ZONE3, clampZone)
        
    Else
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                rgbCombiRefFlatFieldMeanR(site) = 1
                rgbCombiRefFlatFieldMeanG(site) = 1
                rgbCombiRefFlatFieldMeanB(site) = 1
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

Private Sub RGBCombiCM_MakeFlatFieldImagePlane(ByVal pType As String, ByVal ZONE_FULL As Variant, ByVal ZONE_ZONE3 As Variant, ByVal clampZone As Variant)

    '[Edit Here] Filename Prefix Setting
    Const FF_FILE_PREFIX As String = "RGBCombiHL-"
    
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
    Call Average(rRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, rgbCombiRefFlatFieldMeanR)
    Call Average(gRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, rgbCombiRefFlatFieldMeanG)
    Call Average(bRawPlane, "YLINE_ZONE2D", EEE_COLOR_FLAT, rgbCombiRefFlatFieldMeanB)

    'To deposit R/G/B flat field images to the Plane Bank.
    Call TheIDP.PlaneBank.Add("RGBCombiCM FLAT FIELD RED", rRawPlane, True, True)
    Call TheIDP.PlaneBank.Add("RGBCombiCM FLAT FIELD GREEN", gRawPlane, True, True)
    Call TheIDP.PlaneBank.Add("RGBCombiCM FLAT FIELD BLUE", bRawPlane, True, True)

    Call ReleasePlane(rawPlane)

End Sub

'-------------------------------------------------------------------------------------
'# Name
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
Public Sub RGBCombiCM_RGBandIRseparate(InputPlane As CImgPlane, InputZone As Variant, ByRef result() As Double)

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
        Set frShadingPlane = TheIDP.PlaneBank.Item("RGBCombiCM FLAT FIELD RED")
        Set fgShadingPlane = TheIDP.PlaneBank.Item("RGBCombiCM FLAT FIELD GREEN")
        Set fbShadingPlane = TheIDP.PlaneBank.Item("RGBCombiCM FLAT FIELD BLUE")
    
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
        result(RGBCombiRGBColorArray.red, site) = redMean(site)
        result(RGBCombiRGBColorArray.green, site) = greenMean(site)
        result(RGBCombiRGBColorArray.blue, site) = blueMean(site)
    Next site

End Sub

Private Sub ker_RGBCombiCM_v30()

    With TheIDP
        .CreateKernel LowPassFilterNameH, idpKernelFloat, 15, 1, 0, LpfKernel
        .CreateKernel LowPassFilterNameV, idpKernelFloat, 1, 15, 0, LpfKernel
        
        .CreateKernel "kernel_RGBCombiCM3x3", idpKernelInteger, 3, 3, 0, "1 1 1 1 1 1 1 1 1"
'                               1   1   1
'                               1   1   1
'                               1   1   1

        .CreateKernel "kernel_RGBCombiCM5x5", idpKernelInteger, 5, 5, 0, "1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1"
'                               1   1   1   1   1
'                               1   1   1   1   1
'                               1   1   1   1   1
'                               1   1   1   1   1
'                               1   1   1   1   1
        .CreateKernel "kernel_RGBCombiYgloblal2", idpKernelInteger, 9, 7, 0, "1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1 1"
'                             TAP:X9 , Y7
    End With
    
End Sub

Public Sub RGBCombiCM_GetMuraParameters(ByVal paramSheetName As String)
    Dim MuraCol As Collection
    Set MuraCol = New Collection

    Call RGBCombiCM_GetMuraParam(MuraCol, paramSheetName)
    
    Call RGBCombiCM_SetCommonParam(MuraCol)
    Call RGBCombiCM_SetModParam(MuraCol)
    
    Set MuraCol = Nothing
End Sub


Private Sub RGBCombiCM_GetMuraParam(ByRef MuraCol As Collection, ByVal paramSheetName As String)
    
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

Private Function RGBCombiCM_SetCommonParam(ByRef MuraCol As Collection)

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

Public Function RGBCombiCM_SetModParam(ByRef MuraCol As Collection)
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

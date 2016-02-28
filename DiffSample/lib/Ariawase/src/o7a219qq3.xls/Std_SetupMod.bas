Attribute VB_Name = "Std_SetupMod"
Option Explicit
'概要:
'   XEeeAuto_SetupInitTestの下請け関数群
'
'目的:
'   実際の初期化のほとんどはここで行われる
'
'作成者:
'   2012/02/03 Ver0.1 D.Maruyama
'   2012/03/02 Ver0.2 D.Maruyama 中務さん打ち合わせの結果を反映
'                                以下３つの変数がXEeeAuto_Captureへ移動
'                                   Public MIFCObj As New MIFC
'                                   Public eCapState() As CAP_STATE
'                                   Public lngErrorInfo() As Long
'                                InitiTestの引数を削除し、固定マクロ名は内部でもつ
'   2012/04/06 Ver0.3 D.Maruyama Offsetマネージャチェック関数のコールを追加
'   2012/04/09 Ver0.4 D.Maruyama 変数初期化の問題に対応
'                                PTEMPをFirstExec=0の外に出した
'                                SUBのOPENフラグの初期化をCLSVAR,InitTestの2箇所やっていたので、CLSVARのみに変更
'                                IGXLのインターポーズファンクションは別のモジュールとして定義
'   2012/10/18 Ver0.5 H.Arikawa  P7A136LQ3をベースに再作成
'   2012/12/21 Ver0.6 H.Arikawa  LoadRefImageを追加
'   2012/12/26 Ver0.7 H.Arikawa  SRAM/OTPのInitialize部分を再度変更
'   2012/12/26 Ver0.8 H.Arikawa  SetOpen_SiteをDisconnectAllDevicePinsと置き換えの為、削除
'                                GND_Connectを追加、GND_Disconnectを編集。
'   2013/01/07 Ver0.9 H.Arikawa  TENKEN FLAG設定を追加。
'   2013/01/24 Ver1.0 H.Arikawa  OTP Failsafe FLAG設定を追加。
'                                TesterTypeを修正。
'   2013/01/25 Ver1.1 H.Arikawa  OTP Init/SRAM Initをライブラリ吸収に変更。
'   2013/01/31 Ver1.2 H.Arikawa  dc_setupのテストインスタンスより情報取得する処理を削除。
'                                (Phase1で使用していたFW_SeparateFailSiteGndで使用しているが、DCの関数内で代替え済の為、削除)
'   2013/02/01 Ver1.3 H.Arikawa  dc_setupにDcTestLastNumber取得ルーチンを追加。
'   2013/02/04 Ver1.4 H.Arikawa   Flg_LastProcessInfoUseを追加(前工程情報取得)
'   2013/02/20 Ver1.5 H.Arikawa  GND_DisConnect,GND_Connectを修正。
'   2013/02/22 Ver1.6 H.Arikawa  SetKeepAliveを追加。
'   2013/03/01 Ver1.7 H.Arikawa  Flg_Simulator処理追加。OtpInit、SRAMInitの関数(受け場所)作成
'   2013/11/05 Ver1.8 H.Arikawa  TheIDP.Initializeを追加。
'   2013/11/05 Ver1.9 H.Arikawa  TheIDP.Initialize関連の修正。
'
'Tool対応後にコメント外して自動生成にする。　2013/03/07 H.Arikawa
#Const CUB_UB_USE = 0    'CUB UBの設定          0：未使用、0以外：使用
#Const EEE_AUTO_JOB_LOCATE = 2      '1:長崎200mm,2:長崎300mm,3:熊本

Public Const gIsTeradyneDecoder As Boolean = True

Public CurrentJob As String
Public Rga_val As Double
Public Rgb_val As Double
Public Bga_val As Double
Public Bgb_val As Double
Public Gga_val As Double
Public Ggb_val As Double
Public Rga_ref As Double
Public Rgb_ref As Double
Public Bga_ref As Double
Public Bgb_ref As Double
Public Gga_ref As Double
Public Ggb_ref As Double

Public Flg_vsubOpen(nSite) As Boolean

Private SaveWaferNo As String
Public Const TesterType As String = "IP750EX"        '2012/12/20 H.Arikawa Add TesterType情報を取得する変数。

Dim gPDF As PatDriveFormat
Dim gFC As FunctionConstants

Private Sub setFlagsAndNode()

    Dim site As Long
    
    '/* === SETUP DEBUG FLAGS ========================================= */
    '===== Type Common Setup Must Use Item =====
    TheExec.RunOptions.AutoAcquire = True    'TOPT Enable FLAG             True:ON   FALSE:OFF
    Flg_Simulator = 0                        'SIMULATOR MODE FLAG             1:ON       0:OFF
    Flg_Debug = 0                            'DEBUG MODE FLAG                 1:ON       0:OFF
    Flg_Illum_Disable = 0                    'ILLUMINATOR SKIP FLAG           1:ON       0:OFF
    OTPBWC_ERR = 0                           'OTP Failsafe FLAG               1:NG       0:OK
    Flg_LastProcessInfoUse = False           'LastProcessInfo FILE USE FLAG   True:Use   False:NoUse
    Flg_HashCheckResult = True               'HashCode Check Result FLAG      True:OK    False:NG
    EEEAUTO_AUTO_MODIFY_TESTCONDITION = False 'TestConditionAutoOptimize FLAG True:EnableFalse:Disable
    
    '===== Type Common Setup Comment =====
    If Flg_Simulator = 1 Then TheExec.Datalog.WriteComment "Simulator MODE!!"
    If Flg_Illum_Disable = 1 Then TheExec.Datalog.WriteComment "Illuminator Stop MODE!!"
    
    '===== Type Custom Setup Comment =====
    Call TypeCustomFlagSet
        
    '/* =============================================================== */
    
    '最初に撮像項目が全部流れていないと強制敵にAutoAcquireをOFFにする。ムラなどのRef画像読み込みが完了するまでAutoAcquireをOFFする目的。
    'Force disable Auto-Acquire all "image" test has not yet been passed. This is to avoid auto image acquire to run
    'on temporary re-activated sites while reading reference images.(MM)
    If First_Exec = 0 Then Flg_FirstCompleteRun = False
    If Flg_FirstCompleteRun = False Then TheExec.RunOptions.AutoAcquire = False
    
    If TheExec.RunOptions.AutoAcquire = True Then
        TheExec.Datalog.WriteComment "***** Parallel MODE *****"
    Else
        TheExec.Datalog.WriteComment "***** Serial MODE *****"
    End If

'    For Standard library
    If Flg_Simulator = 1 Then
        Flg_Illum_Disable = 1
    End If
    
'    TOPT Setting For TENKEN
    If Flg_Tenken = 1 Then
        TheExec.RunOptions.AutoAcquire = False
        TheExec.RunMode = runModeProduction 'CableCheck
    End If
    
    For site = 0 To nSite
        Flg_FailSite(site) = False
        Flg_FailSiteImage(site) = False
    Next site

End Sub

Public Function dc_setup() As Long

    '### TOPT FW: JobStart時にエラーが発生していたら止める ### 'V210
    If FailedJobInitialize Then
        dc_setup = TL_ERROR
        Exit Function
    End If
    
    '========== TIME MESURE START =========================
    Call StartTime
    
    '========== INIT FOR LastProcessInfo ==================
    If Flg_AutoMode = True And Flg_LastProcessInfoUse = True Then
        If CInt(DeviceNumber_site(0)) = 1 Then  'デバイスNo.が1の時
            Call Init_LastProcessInfoFILE
        End If
    End If
    
    '========== ACTIVE SITE CHECK =========================
    Call mf_ChipExistenceCheck
    
   '======= Settings of the "word" ===================================
    Call FlagSetup
    
    '======= To set debug flags and node id ===========================
    Call setFlagsAndNode
    
    '========== DATALOG ARRANGE PRINT =====================
    If Flg_Print = 1 Then
        Call printMyHeader
    End If
    
    '========== CLEAR =====================
    Call Clsvar
    blnFlg_BlowCheck = False

    '========== INITAL SETUP ==============================
    If First_Exec = 0 Then
        TheHdw.DIB.powerOn = True
        TheHdw.DIB.LeavePowerOn = True
        Call InitJob
        Call JobEnvInit
        #If EEE_AUTO_JOB_LOCATE = 1 Then '長崎200mm
        Call MapOutput
        #End If
            
        Call RVMM_Initialize
        
        '========= CImgIDP 関連の初期化 =========================
        Call TheIDP.Initialize              'TheIDP 関連オブジェクトの生成＆初期化(保存用固定バンク画像含む)
        Set TheParameterBank = Nothing
        Call CreatePlaneMapIfNothing        'PlaneManagerInit設定
        Call CreatePMDIfNothing             'PMD 設定。TheIDP.Initialize で初期化して設定
        Call CreateKernelManagerIfNothing   'カーネル設定。TheIDP.Initialize で初期化して設定
        Call CreateTheParameterBankIfNothing
        
        '========= パターンロード =========================
        If Flg_Simulator = 0 Then Call LoadPatternFile
                
        If Flg_Simulator = 0 Then Call LoadRefImage 'LoadRefImage SheetからRef画像、同期コードRef画像を読み込む。
        
       '========= For Human Error ==========
        If Flg_Simulator = 0 Then Call GetCsvFileName
        If Flg_Simulator = 0 Then Call Get_Hard_data
        If Flg_Simulator = 0 Then Call ReadOffsetFile
        If Flg_Simulator = 0 Then Call WriteOffsetManager

        '===== CAPTURE UNIT INITIALIZE ========================
        If Flg_Simulator = 0 Then Call InitializeCaptureUnitInside
        
        If Flg_Simulator = 0 Then Call OptIni
        
        If Flg_Simulator = 0 Then
            '========= KeepAliveSet ==========
            TheHdw.Digital.Timing.Load
            Call SetKeepAlive
        End If
        
        '===== INIT ReadResponseTime ========================
        If Flg_Simulator = 0 Then Call InitialReadRRT
        
        '========RankSheet Check ==============================
        Call InitializeEeeAutoRank
        
        If IsRankEnable Then
            Call RankInit
        End If
        
        Call GetDeviceType
                
        '========= GET TheExec.CurrentJob Name =======================
        CurrentJobName = TheExec.CurrentJob '2012/12/10 Arikawa Debug
        
        '========= GET DCTestLastNumber ==============================
        DcTestLastNumber = Get_DcTestLastNo()
        
        '========= GET Last Test Number with enable-word is "image"  ==============================(MM)
        ImageTestLastNumber = Get_ImageTestLastNo()
        
        First_Exec = 1
        
        '===== HashCode Check ========================
        If Flg_Simulator = 0 Then
            Call RVMM_GetRegisterVersion
        End If
        
        '===== CSV File FailSafe =====
        If Flg_Simulator = 0 Then Call AllCSVCheckSub
        
        '==== OTP Test Initialize ======================
        Call OtpInit
                
        '==== SRAM Test Initialize ======================
        If Flg_Simulator = 0 Then Call SramInit
        
        '==== OffSet Test Check ======================
        Call CheckAllOffsetExist
        
'        '==== AFE DAC Initialize ====
'        Call Afe_Init(ThisWorkbook)
        
    End If
    
    '==== 1LSB XLsb_NonConversion ====
    Call XLsb_NonConversion_Calc
    
    TenkenTemp = 0
    If Flg_Simulator = 0 Then
        TenkenTemp = 0
        Call ReadTemp(TenkenTemp)
    End If

    
    '### TheIDPのKernel,LUT等のユーザー向け初期化 ################### 'V210
    
    If Not TheIDP.IsExistLUT Then
        Call lut_set
    End If
    
    Call TheIDP.ResetTest   'Flag Clear(ver2.00)
    Call TheVarBank.Clear
    
'    InitTestScenario
   
    InitializeEeeAutoModules 'EeeAutoModの初期化
    
    'FirstExec外のキャプチャユニット初期化
    If Flg_Simulator = 0 Then Call InitializeCaptureUnitOutSide
    
    '========== DEFFECT FILE OPEN =========================
    Call mf_OpenDefectFile
        
    '### Management of ProbeCardContact #############################
    'Kumamoto Only
    #If EEE_AUTO_JOB_LOCATE = 3 Then
        Call Init_ProbeCardInfoFILE
    #End If
    '################################################################

End Function
Private Sub Clsvar()

    Erase Flg_vsubOpen      'sub電圧フラグの初期化
    
    Erase Logic_judge
    
    'Grade
    Erase Ng_test, Watchs, Watcht, Watchc, Now_Time, Now_Day  '2012/11/16 175JobMakeDebug
    Erase S_rank, Rank_ng
    Erase Rank_ng
    Erase G2ngbn, G2_flg, G2rank, Rselect2
    Erase G3ngbn, G3_flg, G3rank, Rselect3
    Erase G4ngbn, G4_flg, G4rank, Rselect4
    Erase G5ngbn, G5_flg, G5rank, Rselect5

End Sub

'*************************************************
'**                                             **
'**     Relay Set                               **
'**                                             **
'*************************************************

Public Sub SET_RELAY_CONDITION(ByVal strApmuUBSet As String, ByVal strCubUBSet As String)

    If strApmuUBSet <> EEE_AUTO_NOUSE_RELAY Then
    
        Call TheUB.AsAPMU.SetUBCondition(strApmuUBSet)
        If IsSnapshotOn = True Then
            Call OutputOptsetInfo("SET_RELAY_CONDITION APMU_UB", TheExec.DataManager.InstanceName, strApmuUBSet)
        End If
    
    End If
    
#If CUB_UB_USE <> 0 Then

    If strCubUBSet <> EEE_AUTO_NOUSE_RELAY Then
    
        Call TheUB.AsCUB.SetUBCondition(strCubUBSet)
    
        If IsSnapshotOn = True Then
            Call OutputOptsetInfo("SET_RELAY_CONDITION CUB_UB", TheExec.DataManager.InstanceName, strCubUBSet)
        End If
    
    End If
    
#End If

End Sub

'============================================================
'     Functions for Job Program
'============================================================

Public Function read_sr() As Boolean
 
    Dim fp As Integer
    Dim fname As String

    On Error GoTo ErrorDetected

    fname = ".\PAR\SystemBoardRef.dat"

    fp = FreeFile
    Open fname For Input As fp

    '====== Kando_hi Keisuu =========
    Input #fp, Rga_ref
    Input #fp, Rgb_ref
    Input #fp, Gga_ref
    Input #fp, Ggb_ref
    Input #fp, Bga_ref
    Input #fp, Bgb_ref
    '================================

    Close fp

    If Sysname = "" Then
        Rga_val = 1
        Rgb_val = 0
        Gga_val = 1
        Ggb_val = 0
        Bga_val = 1
        Bgb_val = 0
    Else
        fname = ".\PAR\" & Sysname & ".dat"

        fp = FreeFile
        Open fname For Input As fp

        '====== Kando_hi Keisuu =========
        Input #fp, Rga_val
        Input #fp, Rgb_val
        Input #fp, Gga_val
        Input #fp, Ggb_val
        Input #fp, Bga_val
        Input #fp, Bgb_val
        '================================

        Close fp
    End If

    read_sr = True
    Exit Function

ErrorDetected:

    outPutMessage "[Error] in read_sr()"
    read_sr = False
    Exit Function

End Function

'$$$$$$$$$$ Save Image DATA for Auto Test $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub GetDeviceType()

    Dim wkshtObj As Object
    
    Const ProductionIFSheetName = "Production IF"

    '======= WorkSheet Select =============
    Set wkshtObj = ThisWorkbook.Sheets(ProductionIFSheetName)
         
    '======= WorkSheet ErrorProcess =======
    If IsEmpty(wkshtObj) Then
        MsgBox "Non Production IF WorkSheet!!"
        Exit Sub
    End If

    '======= Get Device Type ==============
    DeviceType = wkshtObj.Cells(3, 2)

End Sub
Public Sub lut_set()

'{
'   PMD設定、Kernel設定と一緒に管理した方がいい。
'   TheIDP.RemoveResourcesでそれらとまとめてクリアされるので。
'   名前はやりたい内容を表すものに変更。
'}

    Dim intLoopCount As Long
    Dim lngOutVal As Long

'   /****** [1] *****/
    TheIDP.CreateIndexLUT "lut_1", -2048, 2047, 0, 4095, 12              ' Look Up Table 1

'   /****** [2] *****/
    TheIDP.CreateIndexLUT "lut_2", 0, 32767, 0, 32767, 16                 ' Look Up Table 2
    TheIDP.CreateIndexLUT "lut_2", -32767, -1, 32767, 1, 16

End Sub
Private Function ptemp_f() As Double

    Dim site As Long
    Call SiteCheck
    Dim ptemp(nSite) As Double

    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            ptemp(site) = TenkenTemp
        End If
    Next site
    
    Call ResultAdd(GetInstansNameAsUCase, ptemp)
    Call test(ptemp)
    
End Function
Public Sub GND_DisConnect(ByVal targetSite As Long)
    If nSite > 1 Then
        '========== 未使用SITEのGND分離 ===========================================
        Call SET_RELAY_CONDITION("GND_Separate_Site" & targetSite, "-") '2012/11/16 175Debug Arikawa
    End If
                  
End Sub
Public Sub GND_Connect(ByVal targetSite As Long)
    If nSite > 1 Then
        '========== 未使用SITEのGND分離 ===========================================
        Call SET_RELAY_CONDITION("GND_Beta_Site" & targetSite, "-")      '2012/11/16 175Debug Arikawa
    End If
End Sub

Public Sub SetKeepAlive()
    
    With TheHdw.Digital.KeepAlive
        .EraseRAM
        .Count = 1
        .Pins("KeepHiPins").SetRAM 0, "1"
        .Pins("KeepLoPins").SetRAM 0, "0"
    End With
    Call TheHdw.Digital.HRAM.SetCapture(captSTV, False)
    Call TheHdw.Digital.HRAM.SetTrigger(trigFirst, False, 0, True)

End Sub

Public Sub OtpInit()

    Call OtpVariableClear
    Call OtpInitialize_Get_AddressParPage
    Call OtpInitialize_Get_FixedValue
    Call OtpInitialize_Get_FFBlowInfo
    Call OtpInitialize_Select_OtpBlow_Page
    Call OtpInitialize_Make_FixedValuePattern

    'OTPBLOW FLAG SET
    If Flg_AutoMode = True And Flg_Tenken = 0 Then
        Flg_OTP_BLOW = 1
    Else
        Flg_OTP_BLOW = 0
    End If

    Call OtpInitialize_Get_PageBit("Lot1", BitWidthAll_Lot1, Page_Lot1, Bit_Lot1)
    Call OtpInitialize_Get_PageBit("Lot2", BitWidthAll_Lot2, Page_Lot2, Bit_Lot2)
    Call OtpInitialize_Get_PageBit("Lot7", BitWidthAll_Lot7, Page_Lot7, Bit_Lot7)
    Call OtpInitialize_Get_PageBit("Lot8", BitWidthAll_Lot8, Page_Lot8, Bit_Lot8)
    Call OtpInitialize_Get_PageBit("Lot9", BitWidthAll_Lot9, Page_Lot9, Bit_Lot9)
    Call OtpInitialize_Get_PageBit("Wafer", BitWidthAll_Wafer, Page_Wafer, Bit_Wafer)
    Call OtpInitialize_Get_PageBit("Chip", BitWidthAll_Chip, Page_Chip, Bit_Chip)
    Call OtpInitialize_Get_PageBit("Single_CP_FD", BitWidthAll_Single_CP_FD, Page_Single_CP_FD, Bit_Single_CP_FD)
    Call OtpInitialize_Get_PageBit("TEMP", BitWidthAll_TEMP, Page_TEMP, Bit_TEMP)
    Call OtpInitialize_Get_PageBit("SRAM", BitWidthAll_SRAM, Page_SRAM, Bit_SRAM)

'OTPに対するFirst_exec処理

    '■■■■■ 初期設定 ■■■■■


    '■■■■■ 変動値Blow設定 ■■■■■


End Sub

Public Sub SramInit()
'SRAMに対するFirst_exec処理

    '■■■■■ 初期設定 ■■■■■


    Call ValiableSet_SramDesignInfo_IO
    Call ValiableSet_SramDesignInfo_RCON
    
    Call READ_TBL_LIST
    Call READ_TBL_FILE
End Sub
Public Sub XLsb_NonConversion_Calc()

    Dim tmpVar(nSite) As Double
    Dim site As Long
    
    For site = 0 To nSite
        tmpVar(site) = 1
    Next site
    
    Call XLibTheDeviceProfilerUtility.SetLSBParam("XLsb_NonConversion", tmpVar)

End Sub

Public Function GetAfeConstant( _
    sheet_name As String, _
    constant_name As String) As Variant
    
    GetAfeConstant = gFC.GetValue(sheet_name, constant_name)
    
End Function

Public Function GetDacMeasurePin( _
    sheet_name As String, _
    pin_number As Long) As String
    
    If pin_number > 2 Or pin_number < 1 Then
        MsgBox "pin_numberには1か2のみ指定可能です"
    End If
    
    GetDacMeasurePin = gPDF.GetPinName(sheet_name, "OutPIN", pin_number)
    
End Function

Sub Afe_Init(ByRef wbook As Workbook)

On Error GoTo ReleaseObjects
    
    ' PatDriveFormatシートの読み込み
    Set gPDF = Nothing
    Set gPDF = New PatDriveFormat
    Call gPDF.Initialize(wbook)
    
            
    ' FunctionConstantsシートの読み込み
    Set gFC = Nothing
    Set gFC = New FunctionConstants
    Call gFC.Initialize(wbook)
        
ReleaseObjects:
    
End Sub


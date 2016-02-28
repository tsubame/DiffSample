Attribute VB_Name = "CommonEnvModule"
'
'作成者:
'   2012/10/18 Ver0.1 H.Arikawa    Base:p7a136lq3
'   2012/10/18 Ver0.2 H.Arikawa    Tenkenシートより情報取得を行うルーチン追加
'                                  ・GetTenkenSetupInfo
'                                  TenkenのSetupに必要な情報をシートから取得するように変更
'                                  ・SetupTenkenData
'                                  Locationフラグ追加(Phase1から引き継ぎ)
'                                  ・定義部
'   2013/01/24 Ver0.3 H.Arikawa    OTPBWC_ERRを追加(OTPfailsafe Flag)
'   2013/02/04 Ver0.4 H.Arikawa    Flg_LastProcessInfoUseを追加(前工程情報取得)
'                                  ICUL1G_USE、HSD200_USEの条件付きコンパイル引数の定義追加。

Option Explicit
'##########################
'#   EDITABLE VARIABLES   #
'##########################

'======== Job Name Add ================================
Public Const NormalJobName As String = "o7a219qq3"          'Check!!Auto-Insert
'======================================================

'=========== Select Wafer Map =========================
Public Const Map_fname As String = "IMX219QQ_ES_Wafer_o7.map"        'Check!!Auto-Insert
'======================================================

'=========== Select OutputImage Place =================
Public Const gCaptureDirectory As String = "Z:\imx219\"
'======================================================

'======== SITE VARIABLES ================
Public Const SITE_MAX As Integer = 8   'Check!!Auto-Insert
Public Const nSite As Long = SITE_MAX - 1
Public Check_site(nSite) As Integer

Public LimitSetIndex As Integer
Public Sw_Node As Integer
Public Sw_sr As Integer
Public Flg_Shift As Integer
Public Flg_Internal As Integer
Public C_numb As Integer
Public Flg_Debug As Integer
Public Sw_lop As Integer
Public Sw_Tenken As Integer
Public Sw_Ana As Integer
Public Sw_Cbar As Integer
Public First_Exec As Integer
Public LogName As String
Public Flg_Simulator As Integer
Public Flg_Illuminator As Integer
Public Flg_Print As Integer
Public Flg_Scrn As Integer
Public Flg_DacLog As Integer
Public Flg_Capture As Integer
Public Flg_Shmoo As Long
Public Flg_Cnd As Long
Public OTPBWC_ERR As Long           '2013/01/24 H.Arikawa Add OTP FailSafeFlag
Public Flg_LastProcessInfoUse As Boolean  '2013/02/4 H.Arikawa Add
Public Max_Retry As Integer
Public Flg_ImageCapture As Boolean
Public DebugI As Double
Public EEEAUTO_AUTO_MODIFY_TESTCONDITION As Boolean     '2013/10/28 H.Arikawa TestCondition Auto Optimize Flag

Public Flg_cmos As Integer          '02/06/10
Public Flg_image As Integer
Public Flg_color As Integer
Public Flg_margin As Integer        '00/07/25
Public Flg_shiroten As Integer      '00/07/25
Public Flg_Tenken As Integer        '02/08/07
Public Flg_AutoMode As Boolean
Public Flg_c_Command As Boolean
Public Flg_LoopMode As Boolean
Public Apmu_Alarm As Boolean

'+++++++++++++++ Tenken Data Create +++++++++++++++
Public Const Flg_TenkenDataCreate As Boolean = False

'#########################################
'#   RESERVED VARIABLES ( DO NOT EDIT )  #
'#########################################

'###### 条件付きコンパイル設定 ######
#Const EEE_AUTO_JOB_LOCATE = 2      '1:長崎200mm,2:長崎300mm,3:熊本
'===== アサイン関連条件付きコンパイル設定 =====
#Const CUB_UB_USE = 0    'CUB UBの設定          0：未使用、0以外：使用
#Const HDVIS_USE = 0      'HVDISリソースの使用　 0：未使用、0以外：使用  <PALSとEeeAuto両方で使用>
#Const DPS_USE = 1          'DPSリソースの使用　   0：未使用、0以外：使用
#Const APMU_USE = 1        'APMUリソースの使用　  0：未使用、0以外：使用
#Const ICUL1G_USE = 1               '1CUL1Gボードの使用　  0：未使用、0以外：使用  <TesterがIP750EXならDefault:1にしておく>
#Const HSD200_USE = 1               'HSD200ボードの使用　  0：未使用、0以外：使用  <TesterがIP750EXならDefault:1にしておく>

#If EEE_AUTO_JOB_LOCATE = 1 Then     '長崎200mm
Const nodePrefix As String = "SNGST"
#ElseIf EEE_AUTO_JOB_LOCATE = 2 Then '長崎300mm
Const nodePrefix As String = "SNGTX"
Const nodePrefix2 As String = "SKCCDS"  '長崎開発機
#Else '熊本S
Const nodePrefix As String = "SKMBPC"
#End If

'NASAで使用する。
Public Const TenkenJobName As String = "Tenken"

Type TenkenType
    Name As String
    limit As Double
    Value As Double
End Type
Dim TenkenData() As TenkenType

Private Const MaxTenkenSize As Integer = 200

'===== TENKEN sheet =====
Private Const TenkenSheetName As String = "Tenken"

Private Const TenkenCol As Integer = 4         'TenkenシートのTenken項目名が記載してある列番号

Private Const TenkenItemRow As Integer = 2     'TenkenシートのTenkenラベルが記載してある列番号
Private Const TenkenLimitRow As Integer = 3
Private ItemNumber As Long

'===== Add Eee-Auto TENKEN =====
Private Const TenkenValueCol As Integer = 9   'TenkenシートのTenken項目名を記載する開始行番号
Private Const TenkenDefaultNum As Integer = 8   'TenkenシートのTenken項目名を記載する開始行番号
Private Const TenkenCustomNum As Integer = 5    'Default項目以外の項目名を記載する開始行番号
'===== Add Eee-Auto TENKEN =====

'===== control flags =====
Private Const modeAuto As Integer = 0
Private Const modeManual As Integer = 1
Private Const modeLoop As Integer = 2
Private Const modeTenken As Integer = 3
Private Const modeSensRatio As Integer = 4

Private Const condInitial As Integer = 1
Private Const condPreLot As Integer = 2
Private Const condPostLot As Integer = 3
Private Const condPreJob As Integer = 4
Private Const condPostJob As Integer = 5

Private Const stateInitial As Integer = 0
Private Const stateMenu As Integer = 1
Private Const stateRun As Integer = 2

'===== variables for each tests =====
Private TmpPath As String
Private LogFileName As String

Public Chip_f As Integer
Public Chip_x As Integer
Public Chip_y As Integer

Public DeviceNumber_site(nSite)  As Long 'For earch site chip No
Public ChipAdr_x(nSite) As Long
Public ChipAdr_y(nSite) As Long

Public DeviceNumber As Long
Public Defect_full_fname As String
Public LotName As String
Public Workdir As String
Public WaferNo As String
Public Sysname As String
Public Marflg(10) As Integer        '00/07/25 modify
Public DeviceType As String
Public CurrentJobName As String '2012/11/16 175JobMakeDebug Arikawa

'===== declarations for calling windows API =====
Private Declare Function GetComputerName Lib "kernel32.dll" Alias _
    "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
    (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Public StringValue As String
Public CheckValue As Integer
Private Const MasterModuleForEX_VERSION As String = "1.76"

'###########################################################
'#   GET TENKEN LABEL & TENKEN NAME FROM Sheet:Tenken      #
'###########################################################

Public Sub GetTenkenSetupInfo(ByVal RowNum As Integer, ByRef retList() As String, ByRef RetNum As Integer)
    Dim MaxRow As Integer
    Dim StartDataNum As Integer
    Dim i As Integer
    Dim j As Integer
    
'　 内容:
'   JOBのSheet:Tenkenより点検項目数、点検ラベル、点検結果取得項目を取得する。
'   [RowNum]        IN   String型:              リストを取得する列指定
'   [retList()]     OUT  String型:              取得したリスト
    
    MaxRow = Worksheets("Tenken").Range("B4").End(xlDown).Row   'データが入っている最後の最終列を返す
    RetNum = MaxRow - TenkenDefaultNum                                 'TenkenData数を取得

    ReDim retList(RetNum)

    j = 0
    For i = TenkenValueCol To MaxRow
        retList(j) = Sheets("Tenken").Cells(i, RowNum).Value 'TenkenのLabel/Tname名を取得し、Listに格納する(RowNum:2⇒Label,RowNum:4⇒Tname)
        j = j + 1
    Next i
    
End Sub

'###########################################################
'#   EDITABLE FUNCTIONS for TENKEN ( BUT DO NOT REMOVE )   #
'###########################################################

Public Function SetupTenkenData() As Boolean

    Dim site As Long
    Dim TLabel() As String                                          'Add Eee-Auto.2012/10/11 H.Arikawa
    Dim dataNum As Integer                                          'Add Eee-Auto.2012/10/11 H.Arikawa
    Dim i, j As Integer                                             'Add Eee-Auto.2012/10/11 H.Arikawa
    
    On Error GoTo ErrorDetected
    
    Call GetTenkenSetupInfo(TenkenItemRow, TLabel(), dataNum)    'Tenkenラベル名の取得　　　　Add Eee-Auto.2012/10/11 H.Arikawa

    Erase TenkenData
    ReDim TenkenData(MaxTenkenSize)
        
    '***** Please add item names which appear in the tenken sheet from here *****
    TenkenData(0).Name = "WAFR_0"
    TenkenData(1).Name = "XADD_0"
    TenkenData(2).Name = "YADD_0"
    TenkenData(3).Name = "MXLX_0"
    TenkenData(4).Name = "OPTC_0"
    
    i = 0
    j = 0
        
    For site = 0 To nSite
        For i = 0 To (dataNum - 1)
            TenkenData(TenkenCustomNum + j).Name = TLabel(i) & "_" & CStr(site)
            j = j + 1
        Next i
    Next site
    
    SetupTenkenData = True
    Exit Function
    
ErrorDetected:
    outPutMessage "[Error] in SetupTenkenData()"
    SetupTenkenData = False
    Exit Function
    
End Function

Public Function GetTenkenData() As Boolean

    Dim Testval() As Double                                     'Add Eee-Auto.2012/10/11 H.Arikawa
    Dim tname() As String                                       'Add Eee-Auto.2012/10/11 H.Arikawa
    Dim dataNum As Integer                                      'Add Eee-Auto.2012/10/11 H.Arikawa
    Dim i As Integer                                            'Add Eee-Auto.2012/10/11 H.Arikawa
    
    Dim site As Long
    
    On Error GoTo ErrorDetected
    
    Call GetTenkenSetupInfo(TenkenCol, tname(), dataNum)     'Tenken項目名の取得　　　　Add Eee-Auto.2012/10/11 H.Arikawa
        
    '***** Please add actual variables which relate to items in the tenken sheet from here *****
    TenkenData(0).Value = TenkenWaferNo
    TenkenData(1).Value = TenkenX
    TenkenData(2).Value = TenkenY
    TenkenData(3).Value = Opt_Lux
    TenkenData(4).Value = OptResult
    i = 0
    
    'LOGIC項目の測定データを取得する場合は、Test欄に"LOGIC"と記載する。
    For i = 0 To (dataNum - 1)
        If tname(i) = "LOGIC" Then
            For site = 0 To nSite
                Testval(site) = Logic_judge(site)
            Next site
        Else
            TheResult.GetResult tname(i), Testval
        End If
        For site = 0 To nSite
            TenkenData(TenkenCustomNum + i + site * dataNum).Value = Testval(site)
        Next site
    Next i
    
    GetTenkenData = True
    Exit Function
    
ErrorDetected:
    outPutMessage "[Error] in SetupTenkenData()"
    GetTenkenData = False
    Exit Function
    
End Function
'Set/Get values from production control to/from job program
'This function will be called only from production control
Public Function JobInterface(ByVal testmode As Integer, ByVal state As Integer, ByVal Condition As Integer, _
        ByRef IVal() As Integer, ByRef Lval() As Long, ByRef Dval() As Double, ByRef Sval() As String) As Boolean
    Dim i As Integer
    Dim ret As Integer
    Dim str As String
    Dim Length As Integer
    Dim site As Long
    On Error GoTo ErrorDetected

    Select Case Condition
        Case condInitial
            Select Case state
                Case stateInitial
                    'Set log file path and node number
                    JobEnvInit
                    
                    'Set default number
                    LimitSetIndex = 0
                    
                    '##### please set default number #####
                    C_numb = CInt(Mid(NormalJobName, 4, 3))      'Add Eee-Auto.2012/10/16 H.Arikawa
            
                Case stateMenu
'                    If SetChipNumber() = False Then
'                        JobInterface = False
'                        Exit Function
'                    End If
            
                Case stateRun
          
            End Select ' state
        
            'Check job for dummy STDF and definition files
            If tlSearchJobData("", "", "") <> 0 Then
                JobInterface = False
                Exit Function
            End If
        
        
            Case condPreLot
                Select Case state
                    Case stateInitial
                    
                    Case stateMenu
          
                        If testmode = modeTenken Then
                            Flg_Tenken = 1
                        Else
                            Flg_Tenken = 0
                        End If
            
                        If Flg_cmos <> 1 Then
                            Select Case testmode
                                Case modeAuto
                                    Flg_AutoMode = True
                                    Flg_Internal = 0
                              
                                Case modeManual Or modeSensRatio
                                    Flg_Shift = GetCheck("POTENTIAL SHIFT", False, 0)  '2012/11/16 175JobMakeDebug
                                    Select Case Flg_Shift
                                        Case -1
                                          JobInterface = False
                                          Exit Function
                                          
                                        Case 0
                                          Flg_Internal = 0
                                        
                                        Case 1
                                            ret = MsgBox("Already shifted?", vbYesNo + vbQuestion, "POTENTIAL SHIFT")
                                            If ret = vbYes Then
                                                Flg_Internal = 1
                                            Else
                                                Flg_Internal = 0
                                            End If
                                    End Select ' Flg_shift
                              
                            
                              Case modeLoop
                                    Flg_Shift = 0
                                    Flg_Internal = 0
                                    Flg_LoopMode = True
                            
                              Case modeTenken
                                    Flg_Shift = 0
                                    Flg_Internal = 0
                                  
                                    '**** WaferTenken **********
                                    Dim Flg_TestEnd As Boolean
                                    Flg_TestEnd = False
                                    Call TenkenSampleSet(Flg_TestEnd)
                                    If Flg_TestEnd = True Then GoTo ErrorDetected
            
                            End Select ' testmode
                            
                            End If  'Flg_cmos
                            
                    Case stateRun
          
                End Select ' state
        
                Sw_lop = IVal(0)
                Sw_Tenken = IVal(1)
                Sw_sr = IVal(2)
                Flg_Debug = IVal(3)
                Flg_image = IVal(4)
                Flg_margin = IVal(5)
                Flg_shiroten = IVal(6)
                Flg_color = IVal(7)
                Sw_Cbar = IVal(8)
                Sw_Ana = IVal(9)
              
                Workdir = Sval(0)
                LotName = Sval(1)
                Defect_full_fname = Sval(2)
            
'                New_y = 1
        
                Erase Marflg
      
                If Sw_Tenken = 1 Then
                    TheExec.CurrentJob = "TENKEN"
                Else
                    TheExec.CurrentJob = NormalJobName
                End If
        
                If testmode = modeAuto Then
                    If False = InitializeVariable(Flg_AutoMode, Flg_image, Flg_margin, _
                                                    Flg_shiroten, Sw_Ana, LotName, _
                                                    Defect_full_fname, Flg_Shift, _
                                                    LimitSetIndex, C_numb) Then               'Debug Flg tuika hituyo!! 2012/11/12
                        GoTo ErrorDetected
                    End If
                End If

            '++++++++++++++++++++++++ Cable Check ++++++++++++++++++++++++
                If (testmode = modeTenken) And (state = 2) And (Flg_TenkenDataCreate = False) Then
                    TheExec.Validate
                    If CableTenkenCheck() = False Then GoTo ErrorDetected
                End If

            Case condPreJob
                Chip_f = IVal(0)
                Chip_x = IVal(1)
                Chip_y = IVal(2)
                DeviceNumber = Lval(0)
                
                Length = Len(CStr(DeviceNumber))
                If Length = 5 Then
                    WaferNo = Left(DeviceNumber, 1)
                Else
                    WaferNo = Left(DeviceNumber, 2)
                End If
                
                Call MultiLocation_Address(Chip_x, Chip_y, ChipAdr_x, ChipAdr_y)  'Debug Auto refrection

                If testmode = modeAuto Then
                    If False = WriteOrCheckVariable(Flg_AutoMode, Flg_image, Flg_margin, _
                                                    Flg_shiroten, Sw_Ana, LotName, _
                                                    Defect_full_fname, Flg_Shift, _
                                                    LimitSetIndex, C_numb) Then     'Debug Flg tuika hituyou!! 2012/11/12
                        GoTo ErrorDetected
                    End If
                End If
        
            Case condPostJob
    
    End Select ' condition
    
    Erase IVal, Lval, Dval, Sval
    
    Select Case Condition
        Case condInitial
            ReDim IVal(2)
            IVal(0) = Sw_Node
            IVal(1) = C_numb
            
            ReDim Sval(1)
            Sval(0) = Map_fname
      
        Case condPreLot
        
        Case condPostLot
        
        Case condPreJob
        
        Case condPostJob
            If GetTenken() = True Then
                For i = 0 To MaxTenkenSize
                    With TenkenData(i)
                        If .Name = "" Then Exit For
                        
                        ReDim Preserve Dval((i + 1) * 2)
                        ReDim Preserve Sval(i + 1)
                        
                        Sval(i) = .Name
                        Dval(i * 2) = .limit
                        Dval(i * 2 + 1) = .Value
                
                    End With
                Next i
            Else
                JobInterface = False
                Exit Function
            End If
        
    End Select
        
    JobInterface = True
    Exit Function
    
ErrorDetected:
    outPutMessage "Error Detected at JobInterface() in job program ( " & Condition & " )"
    JobInterface = False
    Exit Function

End Function

'###########################################################
'#   MultiLocation_Address For Single                      #
'###########################################################

Public Sub MultiLocation_Address(ByVal Chip_x As Integer, ByVal Chip_y As Integer, ByRef retResult_x() As Long, ByRef retResult_y() As Long)
    Dim site As Long
    
'　 内容:
'   マルチロケーションに対応したアドレス指定のモジュールが挿入される。
'   ロケーションに対応したx,yアドレスを返す。
    
    Dim Loc_X_offset(7) As Integer
Dim Loc_Y_offset(7) As Integer

Loc_X_offset(0) = 0
Loc_Y_offset(0) = 0
Loc_X_offset(1) = -1
Loc_Y_offset(1) = 1
Loc_X_offset(2) = -2
Loc_Y_offset(2) = 2
Loc_X_offset(3) = -3
Loc_Y_offset(3) = 3
Loc_X_offset(4) = -4
Loc_Y_offset(4) = 4
Loc_X_offset(5) = -5
Loc_Y_offset(5) = 5
Loc_X_offset(6) = -6
Loc_Y_offset(6) = 6
Loc_X_offset(7) = -7
Loc_Y_offset(7) = 7

    
    For site = 0 To nSite
        retResult_x(site) = Chip_x + Loc_X_offset(site)
        retResult_y(site) = Chip_y + Loc_Y_offset(site)
    Next site
    
End Sub

'#########################################
'#   RESERVED FUNCTIONS ( DO NOT EDIT )  #
'#########################################
Public Sub JobEnvInit()
    'Set log filename of IG-XL's
    
    If TmpPath = "" Then
        TmpPath = GetTmpPath()
        LogFileName = TmpPath & "errors.txt"
    End If
    Workdir = CurDir

#If Not EEE_AUTO_JOB_LOCATE = 2 Then     '長崎300mm以外
    If Flg_Simulator = 0 Then
        'Get this tester's name
        If Sysname = "" Then
            Sysname = ComputerName()
            If Left$(Sysname, Len(nodePrefix)) = nodePrefix Then
                'Get node number
                Sw_Node = CInt(Right$(Sysname, Len(Sysname) - Len(nodePrefix)))
                outPutMessage "Node number is " & Sw_Node
            Else
                MsgBox "Node name must begin with " & nodePrefix, vbOKOnly + vbExclamation, "Node Name Prefix"
            End If
        End If
    End If
#Else
    If Flg_Simulator = 0 Then
        'Get this tester's name
        If Sysname = "" Then
            Sysname = ComputerName()
            If Left$(Sysname, Len(nodePrefix)) = nodePrefix Then
                'Get node number
                Sw_Node = CInt(Right$(Sysname, Len(Sysname) - Len(nodePrefix)))
                outPutMessage "Node number is " & Sw_Node
            ElseIf Left$(Sysname, Len(nodePrefix2)) = nodePrefix2 Then
                'Get node number
                Sw_Node = CInt(Right$(Sysname, Len(Sysname) - Len(nodePrefix2)))
                outPutMessage "Node number is " & Sw_Node
            Else
                MsgBox "Node name must begin with " & nodePrefix & " or " & nodePrefix2, vbOKOnly + vbExclamation, "Node Name Prefix"
            End If
        End If
    End If
#End If

End Sub

Public Function GetTenken() As Boolean
    Dim wksht As Variant
    Dim i As Integer
    Dim j As Integer
    Dim found As Boolean
    
    Dim tlimit() As String                                       'Add Eee-Auto.2012/10/11 H.Arikawa
    Dim dataNum As Integer                                       'Add Eee-Auto.2012/10/11 H.Arikawa
    Dim site As Long
    On Error GoTo ErrorDetected
    
    GetTenken = False
    If SetupTenkenData() = False Then Exit Function
    If GetTenkenData() = False Then Exit Function
    
    Call GetTenkenSetupInfo(TenkenLimitRow, tlimit(), dataNum)     'Tenken項目名の取得　　　　Add Eee-Auto.2012/10/11 H.Arikawa

    TenkenData(0).limit = 0                                                   'Wafer No.
    TenkenData(1).limit = 0                                                   'X-Address
    TenkenData(2).limit = 0                                                   'Y-Address
    TenkenData(3).limit = Sheets("Tenken").Cells(7, TenkenLimitRow).Value     'OPT MaxLux
    TenkenData(4).limit = 0                                                   'OPT Check
    i = 0
    
    For i = 0 To (dataNum - 1)
        For site = 0 To nSite
            TenkenData(TenkenCustomNum + i + site * dataNum).limit = CDbl(tlimit(i))
        Next site
    Next i
    
    GetTenken = True
    Exit Function
    
ErrorDetected:
    outPutMessage "[Error] in GetTenken()"
    GetTenken = False
    Exit Function
    
End Function

Public Sub outPutMessage(Msg As String, Optional title As String = "Message")
    Dim fp As Integer
    On Error Resume Next

    fp = FreeFile
    Open LogFileName For Append As fp
    Print #fp, Msg
    Close fp
End Sub

Public Function GetTmpPath()
    Dim strFolder As String
    Dim lngResult As Long
    Const MAX_PATH = 128
    On Error Resume Next

    strFolder = String(MAX_PATH, 0)
    lngResult = GetTempPath(MAX_PATH, strFolder)
    If lngResult <> 0 Then
        GetTmpPath = Left(strFolder, InStr(strFolder, Chr(0)) - 1)
    Else
        GetTmpPath = ""
    End If
End Function

Public Function ComputerName() As String
    Dim cn As String
    Dim ls As Long
    Dim res As Long
    On Error Resume Next

    cn = String(1024, 0)
    ls = 1024
    res = GetComputerName(cn, ls)
    If res <> 0 Then
        ComputerName = Mid(cn, 1, InStr(cn, Chr(0)) - 1)
    Else
        ComputerName = ""
    End If
End Function

Public Sub TestFunc()
    Dim status As Long
    status = tlSearchJobData("", "", "")
End Sub
'/*** 17/Sep/02 takayama append ***/
Public Function tlJobChipNo(ByRef ChipNumber() As Integer)
    
    Dim site As Long
    
    For site = 0 To nSite
        DeviceNumber_site(site) = ChipNumber(site)
    Next site
    
End Function

'/*** 11/Mar/02 takayama append
Public Function GetXYAddr(addx As Integer, addy As Integer, Flg_cmos As Integer) As Boolean
    addx = -3
    addy = 0
    Flg_cmos = 1
    
    GetXYAddr = True
End Function

Public Function GetString(Label As String) As String
    On Error Resume Next

    StringValue = ""
    With StringInputForm
        .Setup Label

        .Show
        GetString = StringValue
    End With

End Function

Public Function GetCheck(Label As String, default As Boolean, query As Integer) As Integer  '2012/11/16 175JobMakeDebug
    On Error Resume Next

    CheckValue = -1
    With CheckInputForm
        .Setup Label, default, query

        .Show
        GetCheck = CheckValue
    End With

End Function

Public Function GetTestNumber(ByRef tnum() As Long) As Boolean
    Dim i As Integer
    
    ' for multi-site job
    For i = 0 To nSite
        tnum(i) = CLng(Ng_test(i))   ' copy value
    Next
    
    GetTestNumber = True
End Function

'FloorMonitor MapOutput
Public Sub MapOutput()
    Dim testerName As String
    Dim TesterNo As String
    Dim tname As String
    Dim MapFName As String
    Dim MapOutputDir As String
    On Error Resume Next
    
    If Flg_Simulator = 1 Then Exit Sub
    
    MapOutputDir = "H:\FloorMonitor\MAP\"
    testerName = ComputerName()
             
    If Left$(testerName, Len("SNGST")) = "SNGST" Then
        TesterNo = Right$(testerName, Len(testerName) - Len("SNGST"))
        tname = "ST"
    ElseIf Left$(testerName, Len("SKCCDS")) = "SKCCDS" Then
        TesterNo = Right$(testerName, Len(testerName) - Len("SKCCDS"))
        Select Case TesterNo
         Case 13: TesterNo = "01"
         Case 14: TesterNo = "02"
         Case 15: TesterNo = "03"
         Case Else: TesterNo = "00" 'Dummy
        End Select
        tname = "T#"
    Else
        Exit Sub
    End If
    
    If Len(TesterNo) = 2 Then
        MapFName = tname & "0" & TesterNo & ".MAP"
    Else
        MapFName = tname & TesterNo & ".MAP"
    End If
    
    FileCopy ThisWorkbook.Path & "\par\" & Map_fname, MapOutputDir & MapFName
    
End Sub

Private Function InitializeVariable(ParamArray VariableArr() As Variant) As Boolean
    
    On Error GoTo ErrorInitializeVariable
    InitializeVariable = False

    Const fileName As String = "C:\VariableCheckFile.txt"
    Dim fileNum As Integer
    Dim flag As Boolean
    
    fileNum = FreeFile
    Open fileName For Output As fileNum
    flag = True
    Print #fileNum, "VariableCheckFile INITIALIZE Step1"
    Dim i As Long
    For i = 0 To UBound(VariableArr)
        Print #fileNum, CStr(VariableArr(i))
    Next i
    Close fileNum
    InitializeVariable = True
    Exit Function
    
ErrorInitializeVariable:
    If flag = True Then Close fileNum
    MsgBox "ErrorDetected! @InitializeVariable"
    InitializeVariable = False
End Function

Public Function WriteOrCheckVariable(ParamArray VariableArr() As Variant) As Boolean

    On Error GoTo ErrorWriteOrCheckVariable
    WriteOrCheckVariable = False
    Dim Flg_Pass As Boolean
    Flg_Pass = False

    Const fileName As String = "C:\VariableCheckFile.txt"
    Dim fileNum As Integer
    Dim flag As Boolean
    
    Dim Variable() As Variant
    ReDim Variable(UBound(VariableArr))
    Dim i As Long
    For i = 0 To (UBound(VariableArr))
        Variable(i) = VariableArr(i)
    Next i

    Dim buf As String
    If Dir(fileName) <> "" Then
        fileNum = FreeFile
        Open fileName For Input As fileNum
        flag = True
        Line Input #fileNum, buf
        Close fileNum
        flag = False
        If buf = "VariableCheckFile INITIALIZE Step1" Then
            Flg_Pass = InitVariable(Variable)
        ElseIf buf = "VariableCheckFile INITIALIZE Step2" Then
            Flg_Pass = WriteVariable(Variable)
        Else
            Flg_Pass = CheckVariable(Variable)
        End If
    Else
        MsgBox "Don't Exist Targetfile " & fileName
        WriteOrCheckVariable = False
        Exit Function
    End If
    If Flg_Pass = True Then
        WriteOrCheckVariable = True
    Else
        WriteOrCheckVariable = False
    End If
    Exit Function
    
ErrorWriteOrCheckVariable:
    If flag = True Then Close fileNum
    MsgBox "ErrorDetected! @WriteOrCheckVariable"
    WriteOrCheckVariable = False
End Function

Private Function InitVariable(VariableArr() As Variant) As Boolean
    
    On Error GoTo ErrorInitVariable
    InitVariable = False
    
    Const fileName As String = "C:\VariableCheckFile.txt"
    Dim fileNum As Integer
    Dim flag As Boolean
    
    Dim buf As String
    Dim tmpArr() As Variant
    ReDim tmpArr(UBound(VariableArr))
    
    Dim i As Long, j As Long
    fileNum = FreeFile
    Open fileName For Input As fileNum
    flag = True
    Line Input #fileNum, buf
    i = 0
    Do Until EOF(1)
        Line Input #fileNum, buf
        tmpArr(i) = buf
        i = i + 1
    Loop
    Close fileNum

    fileNum = FreeFile
    Open fileName For Output As fileNum
    flag = True
    Print #fileNum, "VariableCheckFile INITIALIZE Step2"
    For j = 0 To (i - 1)
        Print #fileNum, CStr(tmpArr(j))
    Next j
    Close fileNum
    flag = False
    
    InitVariable = True
    Exit Function
    
ErrorInitVariable:
    If flag = True Then Close fileNum
    MsgBox "ErrorDetected! InitVariable"
    InitVariable = False
End Function

Private Function WriteVariable(VariableArr() As Variant) As Boolean
    
    On Error GoTo ErrorWriteVariable
    WriteVariable = False

    Const fileName As String = "C:\VariableCheckFile.txt"
    Dim fileNum As Integer
    Dim flag As Boolean
    Dim buf As String
    Dim i As Long

    fileNum = FreeFile
    Open fileName For Input As fileNum
    flag = True
    Line Input #fileNum, buf
    i = 0
    Do Until EOF(1)
        Line Input #fileNum, buf
        If CStr(VariableArr(i)) <> buf Then
            MsgBox "Don't Agree  Variable " & fileName
            WriteVariable = False
            Close fileNum
            Exit Function
        End If
        i = i + 1
    Loop
    Close fileNum

    fileNum = FreeFile
    Open fileName For Output As fileNum
    flag = True
    For i = 0 To UBound(VariableArr)
        Print #fileNum, CStr(VariableArr(i))
    Next i
    Close fileNum
    flag = False
    
    WriteVariable = True
    Exit Function
    
ErrorWriteVariable:
    If flag = True Then Close fileNum
    MsgBox "ErrorDetected! @WriteVariable"
    WriteVariable = False
End Function

Private Function CheckVariable(VariableArr() As Variant) As Boolean
    
    On Error GoTo ErrorCheckVariable
    CheckVariable = False
    
    Const fileName As String = "C:\VariableCheckFile.txt"
    Dim fileNum As Integer
    Dim flag As Boolean
    
    Dim buf As String
    Dim i As Long
    fileNum = FreeFile
    Open fileName For Input As fileNum
    flag = True
    i = 0
    Do Until EOF(1)
        Line Input #fileNum, buf
        If CStr(VariableArr(i)) <> buf Then
            MsgBox "Don't Agree  Variable " & fileName
            CheckVariable = False
            Close fileNum
            Exit Function
        End If
        i = i + 1
    Loop
    Close fileNum
    If UBound(VariableArr) <> (i - 1) Then
        MsgBox "Don't Agree  Variable " & fileName
        CheckVariable = False
        Exit Function
    End If
    
    CheckVariable = True
    Exit Function
    
ErrorCheckVariable:
    If flag = True Then Close fileNum
    MsgBox "ErrorDetected! @CheckVariable"
    CheckVariable = False
End Function

Public Function CableTenkenCheck() As Boolean
    
    ItemNumber = 12
    Const TenkenItemSkip As Integer = 5
    Dim i As Integer
    Dim j As Integer
    Dim site As Long
    Dim tmpSite As Long
    Dim SiteMod As Integer
    Dim ChansIndex As Long
    Dim ImageAcqSetsIndex As Long
    Dim CalTdrIOPins As String
    Dim CalTdrICUL1GPins As String
    Dim Flg_FirstColumnIO As Boolean
    Dim Flg_FirstColumnICUL1G As Boolean
    Dim TenkenRefData() As Double
    Dim TenkenLimitData() As Double
    Dim CableTenkenResultData() As Double
    Dim TenkenTestName() As String
    Dim Flg_CableTenkenCheckNG As Boolean
    Dim CableTenkenJudge(nSite) As Boolean
    Dim tmpTenkenDataMax As Integer
    Dim fileNum As Integer
    Dim SingleTenkenResultFile As String
    Dim wksht As Worksheet
    Dim CableCheckMeasureCount As Integer
    Dim CableCheckMeasureSiteJudge As Integer
    Dim CableCheckMeasureSiteContinueMax As Integer
    Dim ActiveSiteJudgeStack(SITE_MAX - 1) As Integer
    Dim ActiveStartSite As Integer

    CableTenkenCheck = False

    If TenkenLimitDataGet(TenkenLimitData, TenkenTestName, tmpTenkenDataMax, TenkenItemSkip) = False Then Exit Function
    If TenkenRefDataGet(TenkenRefData, TenkenItemSkip) = False Then Exit Function

    ReDim Preserve TenkenLimitData(tmpTenkenDataMax - 1)
    ReDim Preserve TenkenTestName(tmpTenkenDataMax - 1)
    ReDim Preserve TenkenRefData(tmpTenkenDataMax - 1)
    ReDim Preserve CableTenkenResultData(tmpTenkenDataMax - 1)

'-------TDRCal I/O ICUL1G PIN READ
    'activate chans sheet
    For Each wksht In ThisWorkbook.Worksheets
        If wksht.Name = "Chans" Then
            wksht.Select
            Exit For
        End If
    Next
    
    If wksht.Name <> "Chans" Then Exit Function

    ChansIndex = 7

    Do While Not IsEmpty(Cells(ChansIndex, 4))
        If Cells(ChansIndex, 4) = "I/O" Then
            If Flg_FirstColumnIO = 0 Then
                CalTdrIOPins = Cells(ChansIndex, 2)
                Flg_FirstColumnIO = 1
            Else
                CalTdrIOPins = CalTdrIOPins & "," & Cells(ChansIndex, 2)
            End If
        End If
        ChansIndex = ChansIndex + 1
    Loop

    'activate Image Acquire Sets sheet
    For Each wksht In ThisWorkbook.Worksheets
        If wksht.Name = "Image_AcqSets" Then
            wksht.Select
            Exit For
        End If
    Next
    
    If wksht.Name <> "Image_AcqSets" Then Exit Function

    ImageAcqSetsIndex = 5

    Do While Not IsEmpty(Cells(ImageAcqSetsIndex, 3))
        If Flg_FirstColumnICUL1G = 0 Then
            CalTdrICUL1GPins = Cells(ImageAcqSetsIndex, 3)
            Flg_FirstColumnICUL1G = 1
        Else
            CalTdrICUL1GPins = CalTdrICUL1GPins & "," & Cells(ImageAcqSetsIndex, 3)
        End If
        ImageAcqSetsIndex = ImageAcqSetsIndex + 1
    Loop

    TheHdw.ICUL1G.ACCalExcludeImgPins CalTdrICUL1GPins
    TheHdw.Digital.ACCalExcludePins CalTdrIOPins
    TheExec.CalibrateTDR
        
    TheExec.RunMode = runModeDebug
    TheExec.RunTestProgram

    CableCheckMeasureCount = Int(WorksheetFunction.Ceiling(Log(SITE_MAX) / Log(2), 1)) - 1
    CableCheckMeasureSiteContinueMax = SITE_MAX
    ActiveStartSite = 0

    SingleTenkenResultFile = "C:\SingleTenkenResultFile.txt"
    fileNum = FreeFile()
    Open SingleTenkenResultFile For Output As #fileNum
    
    If CableCheckMeasureCount > 0 Then
        For j = 0 To CableCheckMeasureCount
            CableCheckMeasureSiteContinueMax = Int(WorksheetFunction.Ceiling(CableCheckMeasureSiteContinueMax / 2, 1)) - 1
            Do While Not site = SITE_MAX - 1
                For site = ActiveStartSite To ActiveStartSite + CableCheckMeasureSiteContinueMax
                    TheExec.sites.site(site).Starting = True
                    ActiveSiteJudgeStack(site) = 1
                    outPutMessage "******** Cable Check Running Site " & site & " ********"
                    If site = SITE_MAX - 1 Then Exit For
                Next
                For site = ActiveStartSite + CableCheckMeasureSiteContinueMax + 1 To ActiveStartSite + CableCheckMeasureSiteContinueMax * 2 + 1
                    TheExec.sites.site(site).Starting = False
                    ActiveSiteJudgeStack(site) = 0
                    If site = SITE_MAX - 1 Then Exit For
                    If site = ActiveStartSite + CableCheckMeasureSiteContinueMax * 2 + 1 Then
                        ActiveStartSite = ActiveStartSite + CableCheckMeasureSiteContinueMax * 2 + 1 + 1
                    End If
                Next
            Loop

            TheExec.RunMode = runModeDebug
            TheExec.RunTestProgram

            For site = 0 To SITE_MAX - 1
                If ActiveSiteJudgeStack(site) = 1 Then
                    Call CableTenkenResultGet(site, CableTenkenResultData)
                End If
            Next

            Print #fileNum, "Site    Result    TestName    Reference    Measure      Limit"
            For i = 0 To tmpTenkenDataMax - 1
                SiteMod = Int(i / ItemNumber)
                If ActiveSiteJudgeStack(SiteMod) = 1 Then
                    If Abs(TenkenRefData(i) - CableTenkenResultData(i)) > TenkenLimitData(i) Then
                        Print #fileNum, " " & SiteMod & "      FAIL      " & TenkenTestName(i) & "_" & SiteMod & "      " & Format(TenkenRefData(i), "0.00E+00") & "     " & Format(CableTenkenResultData(i), "0.00E+00") & "     " & Format(TenkenLimitData(i), "0.00E+00")
                        CableTenkenJudge(SiteMod) = True
                        Flg_CableTenkenCheckNG = True
                    Else
                        Print #fileNum, " " & SiteMod & "      PASS      " & TenkenTestName(i) & "_" & SiteMod & "      " & Format(TenkenRefData(i), "0.00E+00") & "     " & Format(CableTenkenResultData(i), "0.00E+00") & "     " & Format(TenkenLimitData(i), "0.00E+00")
                    End If
                End If
            Next
            Print #fileNum, " "

            For site = 0 To SITE_MAX - 1
                If ActiveSiteJudgeStack(site) = 1 And CableTenkenJudge(site) = True Then
                    outPutMessage "********  Cable Check Site " & site & "  FAIL  ********"
                ElseIf ActiveSiteJudgeStack(site) = 1 And CableTenkenJudge(site) = False Then
                    outPutMessage "********  Cable Check Site " & site & "  PASS  ********"
                End If
            Next
            CableCheckMeasureSiteContinueMax = CableCheckMeasureSiteContinueMax + 1
            ActiveStartSite = 0
        Next
    End If
    
    Close fileNum

    TheExec.RunMode = runModeProduction

    For site = 0 To nSite
        TheExec.sites.site(site).Starting = True
    Next

    If Flg_CableTenkenCheckNG = True Then Exit Function

    CableTenkenCheck = True

End Function
Public Function CableTenkenResultGet(ByVal site As Long, ByRef InCableTenkenResultData() As Double) As Boolean
        
    Dim Testval() As Double                                     'Add Eee-Auto.2012/10/11 H.Arikawa
    Dim tname() As String                                       'Add Eee-Auto.2012/10/11 H.Arikawa
    Dim dataNum As Integer                                      'Add Eee-Auto.2012/10/11 H.Arikawa
    Dim i As Integer                                            'Add Eee-Auto.2012/10/11 H.Arikawa
    
    ReDim Preserve InCableTenkenResultData(1000)
    
    On Error GoTo ErrorCableTenkenResultGet
    
    CableTenkenResultGet = False
    
    Call GetTenkenSetupInfo(TenkenCol, tname(), dataNum)     'Tenken項目名の取得　　　　Add Eee-Auto.2012/10/11 H.Arikawa
        
''''    '***** Please add actual variables which relate to items in the tenken sheet from here *****
'    TenkenData(0).Value = TenkenWaferNo
'    TenkenData(1).Value = TenkenX
'    TenkenData(2).Value = TenkenY
'    TenkenData(3).Value = Opt_Lux
'    TenkenData(4).Value = OptResult
    i = 0
    
    'LOGIC項目の測定データを取得する場合は、Test欄に"LOGIC"と記載する。
    For i = 0 To (dataNum - 1)
        If tname(i) = "LOGIC" Then
'            For site = 0 To nSite
                Testval(site) = Logic_judge(site)
'            Next site
        Else
            TheResult.GetResult tname(i), Testval
        End If
'        For site = 0 To nSite
            InCableTenkenResultData(i + site * dataNum) = Testval(site)
'        Next site
    Next i
    
    CableTenkenResultGet = True
    Exit Function
    
ErrorCableTenkenResultGet:
    outPutMessage "ErrorDetected! @CableTenkenResultGet"
    Exit Function

End Function

Public Function TenkenRefDataGet(ByRef InTenkenRefData() As Double, ByVal InTenkenItemSkip As Integer) As Boolean

    Dim i As Integer
    Dim fileNum As Integer
    Dim tempdata() As String
    ReDim tempdata(1000)
    ReDim InTenkenRefData(1000)
    Dim TenkenRefFileName As String

    On Error GoTo ErrorTenkenRefDataGet

    TenkenRefDataGet = False

'    TenkenRefFileName = ThisWorkbook.Path & "\TENKEN\tenken_ref" & ".dat"
    TenkenRefFileName = ThisWorkbook.Path & "\TENKEN\tenken_ref_" & Format(Sw_Node, "000") & ".dat"


    fileNum = FreeFile()
    Open TenkenRefFileName For Input Access Read As #fileNum

    i = 0
    Do Until EOF(fileNum)
        Line Input #fileNum, tempdata(i)
        If i >= InTenkenItemSkip Then
            InTenkenRefData(i - InTenkenItemSkip) = CDbl(tempdata(i))
        End If
        i = i + 1
    Loop
    
    Close #fileNum
    
    TenkenRefDataGet = True
    Exit Function

ErrorTenkenRefDataGet:
    Close fileNum
    outPutMessage "ErrorDetected! @TenkenRefDataGet"
    Exit Function

End Function

Public Function TenkenLimitDataGet(ByRef InTenkenLimitData() As Double, ByRef InTenkenTestName() As String, ByRef InTenkenDataMax As Integer, ByVal InTenkenItemSkip As Integer) As Boolean
    
    Dim wksht As Worksheet
    Dim site As Long
    Dim i As Integer
    ReDim InTenkenLimitData(MaxTenkenSize)
    ReDim InTenkenTestName(MaxTenkenSize)

    On Error GoTo ErrorTenkenLimitDataGet
    
    TenkenLimitDataGet = False
    
    ' activate tenken sheet
    For Each wksht In ThisWorkbook.Worksheets
        If wksht.Name = TenkenSheetName Then
            wksht.Select
            Exit For
        End If
    Next
    
    If wksht.Name <> TenkenSheetName Then Exit Function
    
    For site = 0 To nSite
        For i = TenkenCol + InTenkenItemSkip To TenkenCol + InTenkenItemSkip + MaxTenkenSize
            If IsEmpty(Cells(i, TenkenItemRow)) Then Exit For
            InTenkenDataMax = InTenkenDataMax + 1
            InTenkenTestName(ItemNumber * site + i - TenkenCol - InTenkenItemSkip) = wksht.Cells(i, TenkenItemRow)
            InTenkenLimitData(ItemNumber * site + i - TenkenCol - InTenkenItemSkip) = CDbl(wksht.Cells(i, TenkenLimitRow))
        Next i
    Next site

    TenkenLimitDataGet = True
    
    Exit Function

ErrorTenkenLimitDataGet:
    outPutMessage "ErrorDetected! @ErrorTenkenLimitDataGet"
    Exit Function

End Function


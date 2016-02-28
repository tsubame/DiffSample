Attribute VB_Name = "DumpREGandRRT_sng"
Option Explicit

' --- Windows API
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    ByVal y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const MAXFILESIZE As Long = 10485760 '10MB
Private Const MAXFILECNT As Long = 3

Declare Function FindWindow Lib "USER32.dll" Alias "FindWindowA" (ByVal lpClassName As String, _
                                                                  ByVal lpWindowName As String) As Long

' --- IG-XL Register APIs
Private Declare Function CALCUB_CALCUB_IDPROM_R Lib "calcub_idl.dll" Alias "_calcub_calcub_idprom_r@0" () As Long
Private Const PCIT_TCI_S0_SETUP_REG As Long = &H148000

Private Declare Function ACOM_SEL_PIN_W Lib "acom_idl.dll" Alias "_acom_sel_pin_w@4" (ByVal Data As Long) As Long
Private Declare Function DILBERT_BOARD_LOC_W Lib "dilbert_idl.dll" Alias "_dilbert_board_loc_w@4" (ByVal Data As Long) As Long
Private Declare Function DILBERT_CONN_STATUS_R Lib "dilbert_idl.dll" Alias "_dilbert_conn_status_r@0" () As Long
Private Declare Function DILBERT_TCTL_READ_CONFIG_R Lib "dilbert_idl.dll" Alias "_dilbert_tctl_read_config_r@0" () As Long

Private Declare Function APMU_MISC_BOARD_SEL_W Lib "apmu_misc_idl.dll" Alias "_apmu_misc_board_sel_w@4" (ByVal Data As Long) As Long
Private Declare Function APMU_MISC_RRT_R Lib "apmu_misc_idl.dll" Alias "_apmu_misc_rrt_r@0" () As Long
Private Declare Function APMU_MISC_MISC0_REG_R Lib "apmu_misc_idl.dll" Alias "_apmu_misc_misc0_reg_r@0" () As Long

Private Declare Function ICUL1G_MISC_BOARD_SEL_W Lib "icul1g_misc_idl.dll" Alias "_icul1g_misc_board_sel_w@4" (ByVal Data As Long) As Long
Private Declare Function ICUL1G_MISC_FPGA_REV_R Lib "icul1g_misc_idl.dll" Alias "_icul1g_misc_fpga_rev_r@0" () As Long

Private Declare Function CTO_BOARD_LOC_W Lib "cto_idl.dll" Alias "_cto_board_loc_w@4" (ByVal Data As Long) As Long
Private Declare Function CTO_GAIN_SRC_R Lib "cto_idl.dll" Alias "_cto_gain_src_r@0" () As Long

Private Declare Function CALCUB_PREC_DAC_DATA_R Lib "calcub_idl.dll" Alias "_calcub_prec_dac_data_r@0" () As Long

Private Declare Function DPS_CHAN_MATCH_W Lib "dps_idl.dll" Alias "_dps_chan_match_w@4" (ByVal Data As Long) As Long
Private Declare Function DPS_RRT_R Lib "dps_idl.dll" Alias "_dps_rrt_r@0" () As Long

Private Declare Function tl_CUBInit Lib "cub.dll" Alias "_tl_CUBInit@4" (ByVal readCalVals As Long) As Long
Private Declare Function tl_CUBReset Lib "cub.dll" Alias "_tl_CUBReset@4" (ByVal powerOn As Long) As Long

Private Declare Function PG_SELECT_W Lib "pg_idl.dll" Alias "_pg_select_w@4" (ByVal Data As Long) As Long
Private Declare Function PG_HDW_INIT_W Lib "pg_idl.dll" Alias "_pg_hdw_init_w@4" (ByVal Data As Long) As Long
Private Declare Function PG_PRIMEDLY_W Lib "pg_idl.dll" Alias "_pg_primedly_w@4" (ByVal Data As Long) As Long

' --- Module Global Variables
Public mlngRefRRMVal As Long
Public mlngRefRRTVal As Long

#Const EEE_AUTO_JOB_LOCATE = 2      '1:í∑çË200mm,2:í∑çË300mm,3:åFñ{
#If EEE_AUTO_JOB_LOCATE = 1 Or EEE_AUTO_JOB_LOCATE = 2 Then     'í∑çË200mm
Private Const DUMPFILEPATH = "G:\jobs\RRT_LOG\DumpRegAndRRT.log"    'í∑çËóp
#Else 'åFñ{S
Private Const DUMPFILEPATH = "F:\Job\CIS\2PC\Debug_Fol\DumpRegAndRRT.log"    'åFñ{óp
#End If


Public Sub InitialReadRRT()

    Dim lngRegVal           As Long
    Dim lngSuperFiledVal    As Long

    ' --- Ping CAL CUB RRT Value in order to measure RRT
    lngRegVal = CALCUB_CALCUB_IDPROM_R

    ' --- Retrive Superfiled Value which contains RRT Measurement Value (called RRM)
    lngSuperFiledVal = tl_bif_r(PCIT_TCI_S0_SETUP_REG)
    
    ' *** For Debug
    'lngRegVal = 10
    'lngSuperFiledVal = 203411524

    ' --- Shift 21 bit since RRM is stored from 21bit to 28bit
    mlngRefRRMVal = tl_shiftn(lngSuperFiledVal, -21)
    
    ' --- Get RRT Value (from 0bit to 7bit)
    mlngRefRRTVal = lngSuperFiledVal And &HFF&

End Sub


Public Function MonitorRRT(Optional blnForcedEnding As Boolean = True, Optional blnRecoverRRT As Boolean = False) As Boolean

    Dim lngSuperFiledVal    As Long
    Dim lngCurrentRRMVal    As Long
    Dim lngCurrentRRTVal    As Long
    Dim lngNewRRTVal        As Long
    Dim lngRegVal           As Long
    Dim blnFlagDataStored   As Boolean
    Dim lngLoopCount        As Long
    Dim lngSiteIndex        As Long

    '-----------------------------------------------------------------------
    blnFlagDataStored = False
    lngLoopCount = 0
    MonitorRRT = True
    
    Do

        ' --- Ping CAL CUB RRT Value in order to measure RRT
        lngRegVal = CALCUB_CALCUB_IDPROM_R

        ' --- Retrive Superfiled Value which contains RRT Measurement Value (called RRM)
        lngSuperFiledVal = tl_bif_r(PCIT_TCI_S0_SETUP_REG)
        
        ' *** For Debug
        'lngRegVal = 10
        'lngSuperFiledVal = 203411524
'        lngSuperFiledVal = 197120068

        ' --- Shift 21 bit since RRM is stored (from 21bit to 28bit)
        lngCurrentRRMVal = tl_shiftn(lngSuperFiledVal, -21)

        ' --- Get RRT Value (from 0bit to 7bit)
        lngCurrentRRTVal = lngSuperFiledVal And &HFF&
        
        If lngCurrentRRMVal <> 0 Then
            blnFlagDataStored = True
            Exit Do
        End If

        lngLoopCount = lngLoopCount + 1

    Loop While lngLoopCount < 5

    If blnFlagDataStored Then

        ' --- Compare Reference RRM Values
        If lngCurrentRRMVal <> mlngRefRRMVal Then

            lngNewRRTVal = lngCurrentRRTVal + (lngCurrentRRMVal - mlngRefRRMVal)

            ' --- Logged When this error has happened
            Call UpdateErrorHistory(mlngRefRRMVal, lngCurrentRRMVal, mlngRefRRTVal, lngCurrentRRTVal)

            ' --- Recover by Adjusting RRT Value
            If blnRecoverRRT = True Then

                ' --- Check if the adjustment value is within 8bit value
                If lngNewRRTVal >= 0 And lngNewRRTVal <= 255 Then

                    ' --- Adjust RRT Value
                    Call tl_bif_w(PCIT_TCI_S0_SETUP_REG, (lngSuperFiledVal And &HFFFFFF00) Or lngNewRRTVal)

                    ' --- Overwrite Reference Value to Adjusted
                    mlngRefRRMVal = lngCurrentRRMVal

                End If
            
            End If
            
            ' --- Forced ending of wafer test
            If blnForcedEnding = True Then
                
                ' --- Save current run mode
                Dim enmCurrentRunMode As RunModeType
                enmCurrentRunMode = TheExec.RunMode
    
                ' --- Change RunMode to DebugMode in order to pop up user form
                TheExec.RunMode = runModeDebug
            
                ' --- Show Error Message to Operator and Flow Error rise
                MsgBox "RRT Error has Occured!!" & vbCrLf & "Re-starting IG-XL should be needed to recover this failure."
'                Err.Raise 999
                MonitorRRT = False
                
                ' --- Back to previous run mode
                TheExec.RunMode = enmCurrentRunMode
                
            End If

        End If

    End If

End Function


Private Sub UpdateErrorHistory(ByVal rrmorg As Long, ByVal rrmerr As Long, ByVal rrtorg As Long, ByVal rrterr As Long)

    Dim filePath As String
    Dim fn As Integer

'    filepath = DUMPFILEPATH & "TCIO_RRT_Data.log"
    filePath = DUMPFILEPATH
    
    RefreshDumpFile
    
    fn = FreeFile
    Open filePath For Append As #fn
    
    Print #fn, "## RRT Info ## " & Format(Now, "yyyy/mm/dd hh:mm:ss") & " System:" & TheHdw.Computer.Name & " Test:" & TheExec.DataManager.InstanceName
    Print #fn, "  Original RRT = " & rrmorg & "(" & rrtorg & "),  Error RRT = " & rrmerr & "(" & rrterr & ")"
    
    Close #fn
    
End Sub


Public Sub dumpPPMUreg()
    
    Dim RVal As Long
    Dim filePath As String
    Dim fw As Integer
    
'    filepath = DUMPFILEPATH & "INSTRUMENT_Reg.log"
    filePath = DUMPFILEPATH
    
    RefreshDumpFile
    
    fw = FreeFile
    Open filePath For Append As #fw
    Print #fw, "## REG Info ## " & Format(Now, "yyyy/mm/dd hh:mm:ss") & " System:" & TheHdw.Computer.Name & " Test:" & TheExec.DataManager.InstanceName
    
    Call ACOM_SEL_PIN_W(TL_PGID_ALL_CHANNELS)
    Call DILBERT_BOARD_LOC_W(&HFFFF)
    RVal = DILBERT_CONN_STATUS_R
    Print #fw, "  CONN_STATUS(ALL) = " & Hex(RVal)
    
    Call DILBERT_BOARD_LOC_W(1)
    RVal = DILBERT_CONN_STATUS_R
    Print #fw, "  CONN_STATUS(0) = " & Hex(RVal)
    RVal = DILBERT_TCTL_READ_CONFIG_R
    Print #fw, "  READ_CONFIG(0) = " & Hex(RVal)
    If RVal <> &HBE Then
        RVal = tl_bif_r(PCIT_TCI_S0_SETUP_REG)
        RVal = tl_shiftn(RVal, -21)
        Print #fw, "  SETUPREG(0)    = " & Hex(RVal)
    End If
    
    Call DILBERT_BOARD_LOC_W(2)
    RVal = DILBERT_CONN_STATUS_R
    Print #fw, "  CONN_STATUS(1) = " & Hex(RVal)
    RVal = DILBERT_TCTL_READ_CONFIG_R
    Print #fw, "  READ_CONFIG(1) = " & Hex(RVal)
    If RVal <> &HBE Then
        RVal = tl_bif_r(PCIT_TCI_S0_SETUP_REG)
        RVal = tl_shiftn(RVal, -21)
        Print #fw, "  SETUPREG(1)    = " & Hex(RVal)
    End If
    
    Call DILBERT_BOARD_LOC_W(&H40)
    RVal = DILBERT_CONN_STATUS_R
    Print #fw, "  CONN_STATUS(6) = " & Hex(RVal)
    RVal = DILBERT_TCTL_READ_CONFIG_R
    Print #fw, "  READ_CONFIG(6) = " & Hex(RVal)
    If RVal <> &HBE Then
        RVal = tl_bif_r(PCIT_TCI_S0_SETUP_REG)
        RVal = tl_shiftn(RVal, -21)
        Print #fw, "  SETUPREG(6)    = " & Hex(RVal)
    End If
    
    Call DILBERT_BOARD_LOC_W(&H80)
    RVal = DILBERT_CONN_STATUS_R
    Print #fw, "  CONN_STATUS(7) = " & Hex(RVal)
    RVal = DILBERT_TCTL_READ_CONFIG_R
    Print #fw, "  READ_CONFIG(7) = " & Hex(RVal)
    If RVal <> &HBE Then
        RVal = tl_bif_r(PCIT_TCI_S0_SETUP_REG)
        RVal = tl_shiftn(RVal, -21)
        Print #fw, "  SETUPREG(7)    = " & Hex(RVal)
    End If
    
    Call APMU_MISC_BOARD_SEL_W(&H20)
    RVal = APMU_MISC_MISC0_REG_R
    Print #fw, "  MISC0(5)       = " & Hex(RVal)
    RVal = APMU_MISC_RRT_R
    Print #fw, "  RRT(5)         = " & Hex(RVal)
    
    Call ICUL1G_MISC_BOARD_SEL_W(&H10000)
    RVal = ICUL1G_MISC_FPGA_REV_R
    Print #fw, "  FPGA_REV(16)   = " & Hex(RVal)
    
    Call CTO_BOARD_LOC_W(&H20000)
    RVal = CTO_GAIN_SRC_R
    Print #fw, "  GAIN_SRC(16)   = " & Hex(RVal)
    
    RVal = CALCUB_PREC_DAC_DATA_R
    Print #fw, "  PREC_DAC(17)   = " & Hex(RVal)
    
    Call ICUL1G_MISC_BOARD_SEL_W(&H80000)
    RVal = ICUL1G_MISC_FPGA_REV_R
    Print #fw, "  FPGA_REV(19)   = " & Hex(RVal)
    
    Call DPS_CHAN_MATCH_W(22)
    RVal = DPS_RRT_R
    Print #fw, "  RRT(22)        = " & Hex(RVal)
    
    Call DPS_CHAN_MATCH_W(23)
    RVal = DPS_RRT_R
    Print #fw, "  RRT(23)        = " & Hex(RVal)
    
    Call ACOM_SEL_PIN_W(TL_PGID_ALL_CHANNELS)
    Call DILBERT_BOARD_LOC_W(&HFFFF)
    RVal = DILBERT_CONN_STATUS_R
    Print #fw, "  CONN_STATUS(ALL) = " & Hex(RVal)
    
    
    If (RVal And &H19) <> 0 Then
        Print #fw, "  #### Resetting CULCUB"
        Close #fw
        fw = FreeFile
        Open filePath For Append As #fw
        Call tl_CUBInit(0)
        Call tl_CUBReset(0)
        
        Print #fw, "  #### Done"
        
        Call ACOM_SEL_PIN_W(TL_PGID_ALL_CHANNELS)
        Call DILBERT_BOARD_LOC_W(&HFFFF)
        RVal = DILBERT_CONN_STATUS_R
        Print #fw, "  CONN_STATUS(ALL) = " & Hex(RVal)
        
        Print #fw, "  #### Init Pipeline"
        ACOM_SEL_PIN_W (1152)
        PG_SELECT_W (2 ^ 18)
        PG_HDW_INIT_W (0)
        PG_PRIMEDLY_W (18)
        PG_HDW_INIT_W (0)
        Call tl_Wait(0.2)
        PG_HDW_INIT_W (1)
        Call tl_Wait(0.2)
        PG_HDW_INIT_W (0)
        PG_HDW_INIT_W (0)
        
        Call ACOM_SEL_PIN_W(TL_PGID_ALL_CHANNELS)
        Call DILBERT_BOARD_LOC_W(&HFFFF)
        RVal = DILBERT_CONN_STATUS_R
        Print #fw, "  CONN_STATUS(ALL) = " & Hex(RVal)
        
        Print #fw, "  #### Reset TCIO"
        Close #fw
        fw = FreeFile
        Open filePath For Append As #fw
        Call tl_tcio_reset
        
        Call ACOM_SEL_PIN_W(TL_PGID_ALL_CHANNELS)
        Call DILBERT_BOARD_LOC_W(&HFFFF)
        RVal = DILBERT_CONN_STATUS_R
        Print #fw, "  CONN_STATUS(ALL) = " & Hex(RVal)
        
    End If
    
    Close #fw
    
    MonitorRRT False
    
End Sub


Private Sub RefreshDumpFile()

    Dim lngCnt As Long
    Dim strFnam As String
    Dim strFnam_sr As String
    Dim strFnam_de As String
    
    If VBA.Dir(DUMPFILEPATH, vbNormal) <> "" Then
        
        If VBA.FileLen(DUMPFILEPATH) > MAXFILESIZE Then
            
            strFnam = Left(DUMPFILEPATH, VBA.InStrRev(DUMPFILEPATH, ".") - 1)
            
            For lngCnt = MAXFILECNT - 1 To 0 Step -1
                    
                If lngCnt > 0 Then
                    strFnam_sr = strFnam & "_" & Trim(str(lngCnt)) & ".log"
                    strFnam_de = strFnam & "_" & Trim(str(lngCnt + 1)) & ".log"
                Else
                    strFnam_sr = DUMPFILEPATH
                    strFnam_de = strFnam & "_" & Trim(str(lngCnt + 1)) & ".log"
                End If
                If VBA.Dir(strFnam_sr, vbNormal) <> "" Then
                    FileCopy strFnam_sr, strFnam_de
                End If
                
            Next lngCnt
            
            Kill DUMPFILEPATH
            
        End If
        
    End If

End Sub


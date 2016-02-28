Attribute VB_Name = "PALS_VoltAdj_Mod"
Option Explicit
'09/09/07           Release                 M.Imamura
'09/09/23           HDVIS/PPMU add.         M.Imamura
'09/11/16           conformed to Eee-JOB    M.Imamura
'09/12/01           DPS add.                M.Imamura
'10/03/04           LimitAlerm (-) add.
'                   Offset >1V regist
'                   PinCheckError->cellcolorYellow
'                   Pincheck Break add.
'                   DataLog Modfied
'                   Cnum/Pnum mod.          M.Imamura
'10/03/15           progressbar del.
'                   site->site_loop
'                   PinCheckError->cellcolorYellow
'                   Break operation mod.    M.Imamura
'10/05/27           impose Check add.
'                   Get_power_condition mod.
'                   ->for Pinname[P_] less  M.Imamura
'10/06/11           HDVIS 1digits taiou     M.Imamura
'10/10/29           ->PALS menber           M.Imamura
'11/10/04  2.30     KeepCondition Add.      M.Imamura
'=======================================
'=======================================
'=============For Production============
'=======================================
'Please insert this command in [StartOfTest] function

'''''If First_Exec = 0 Or Flg_voltsheet_read = 0 Then Call Get_Power_Condition

'=======================================
'============= For Voltcheck============
'=======================================
'immidiate this command at CapturePoint
'Call runpals
'=======================================

'Declare Sub mSecSleep Lib "kernel32" Alias "Sleep" (ByVal lngmSec As Long)

Public Const VOLTTOOLNAME As String = "PALS - Auto Volt Adjust"
Public Const VOLTTOOLVER As String = "2.40"

''Const User Change ===================
Public Const Flag_debug_Voltage = 0 'For debug VoltCheck

Const ResetWaitTime = 1000                                  'Wait For measure(msec)

''=====================================
Const Cha_start_cell_x = 2                           'Start Cell X Adress For ChannelMap
Const Cha_start_cell_y = 7                           'Start Cell Y Adress For ChannelMap

Const Start_x = 2                                    'Condition Cell location on Power-Supply Voltage
Const Start_y = 5                                    'Condition Cell location on Power-Supply Voltage

''GPIB ================================
Const AD7461A_adr = 10                  'Device Address
Dim AD7461A As Integer                  'Device/Discripter

Const ADTIMEOUT = 10
Const ADboard = 0

Dim sts As Integer                      'GPIB I/F Status
Dim cmdstr As String * 40               'GPIB SendData
Dim ADstatus As Integer                 'Status

Dim GetData As String * 40              'GPIB RecieveData

''=====================================

''File ================================
Public Force_Voltage_vc() As Double         'Force_Voltage_vc(condition,pin)
Public Clamp_Current_vc() As Double         'Clamp_Current_vc(condition,pin)
Public Offset_Voltage_vc() As Double        'Offset_Voltage_vc(condition,pin,site)
Dim CheckPin_flag() As Integer              'CheckPin_flag(condition,pin,site)

Public Now_Mode As String                   'Mode of pppset
Public Test_Condition() As String           'TEST CONDITION
Public Cnum As Collection                   'TEST CONDITION
Public Condition_Number As Integer          'CONDITION No
Public Tester_Name() As Integer             'TesterName
Public Tester_exist_cnt As Integer          'TesterCount
Public Site_Number() As Integer             'SiteCount
Public pin_number As Long                   'PIN

Public PowerPinName() As String             'PinName
Public PowerPinNum() As String              'PinNum of CheckPin
Public Pin_Resource() As String             'Pin_Resource from Channel_Map
Public Pin_Chans() As Integer              'Pin_Chans    from Channel_Map
Public CheckRun() As String                 'Do Check Flag of each Pin
Public PNum As Collection
Public Next_Tester_YAddress() As Integer    'Yaddress of sheet
Public Next_Condition_YAddress As Integer   'Yaddress of sheet
''=====================================

''Flag ================================
Public Flg_voltsheet_read As Integer        '1->already Read form Power-Supply Voltage
Public Flg_MeasSkip As Integer
Public Flg_Clear_voltcheck As Integer
Public Flg_Stop_voltcheck As Integer
Public Flg_Break_voltcheck As Integer
Public Flg_End_voltcheck As Integer

Public Flg_SiteLock As Integer
Public Flg_ConditionLock As Integer
Public Flg_PinLock As Integer

Public Flg_ConfirmMode As Integer
Public Flg_Onlyvoltagecheck As Integer
Public Flg_DataLog As Integer
''=====================================

Public Break_Counter As Integer
Public Break_Counter_prev As Integer

Public Now_Site As Long
Public Now_Condition As Integer
Public Now_PinName As Integer

Dim SetupVolt As Double
Dim InputVolt As Double
Dim MeasVolt As Double
Dim DiffVolt As Double
Dim OffsetVolt As Double

Public LimitOffset As Double
Public AdjustOffset As Double

Dim msgtext As String
Dim ANS As Integer

Public FileAddress As Variant        'File Name For Datalog
Public FileNo_vc As Integer

#Const HDVIS_USE = 0      'HVDISリソースの使用　 0：未使用、0以外：使用  <PALSとEeeAuto両方で使用>
#Const DPS_USE = 1          'DPSリソースの使用　   0：未使用、0以外：使用

#Const APMU_USE = 1        'APMUリソースの使用　  0：未使用、0以外：使用
#Const HSD200_USE = 1               'HSD200ボードの使用　  0：未使用、0以外：使用  <TesterがIP750EXならDefault:1にしておく>
#Const BPMU_USE = 1
#Const PPMU_USE = 1


Public Sub sub_VoltFrmShow()
    frm_PALS_VoltAdj_Main.Show
End Sub

Public Sub VoltCheck_start()

    On Error GoTo errPALSVoltCheck_start
    
    LimitOffset = 0.06       'Unit=[V]
    AdjustOffset = 0.003     'Unit=[V]
    
'START VoltCheck Process
    If Flag_debug_Voltage = 1 Then
'     Sw_Node = 71
     Now_Mode = 5
    End If
    
    If Flag_debug_Voltage <> 1 And g_ErrorFlg_PALS = False Then Call Check_GPIB_AD7461A

    If g_ErrorFlg_PALS = False Then
        Worksheets(ReadSheetName).Activate
        Call Get_Power_Condition
        Call Print_Info
        
        If g_ErrorFlg_PALS = False Then
            frm_PALS_VoltAdj_Main.Show
        End If
    End If
    
    Exit Sub

errPALSVoltCheck_start:
    Call sub_errPALS("VoltAdjust Tool Run error at 'VoltCheck_start'", "1-2-01-0-01")

End Sub

Public Sub Check_GPIB_AD7461A()
'Check GPIB board

    Dim Prompt As String

Retry:

    AD7461A = ildev(ADboard, AD7461A_adr, 0, ADTIMEOUT, 1, 1)   ' Send IFC for Initialize
    sts = illn(AD7461A, AD7461A_adr, ALL_SAD, ADstatus)         ' Check Device
    If ADstatus = False Then
        Prompt = "Can't access DIGITAL MULTIMETER" & vbCrLf & _
                "Check  Power  or  Connection GPIB cable"
        ADstatus = MsgBox(Prompt & vbCrLf & "ErrCode.1-2-02-9-02", vbRetryCancel + vbCritical, "Error Masage[VoltCheck]")
        If ADstatus = vbRetry Then
            GoTo Retry
        End If
        If ADstatus = vbCancel Then
            Call sub_errPALS("VoltAdjust Tool Run error at 'Check_GPIB_AD7461A'", "1-2-02-9-03")
            Exit Sub
        End If
    End If

    'Initialize DIGITAL MULTIMETER
    cmdstr = "*RST"                                     'Reset
    sts = ilwrt(AD7461A, cmdstr + vbCrLf, Len(cmdstr))
    Call mSecSleep(500)
    cmdstr = "H0"                                       'Header OFF
    sts = ilwrt(AD7461A, cmdstr + vbCrLf, Len(cmdstr))
    Call mSecSleep(500)

End Sub

Public Sub EndGPIB_AD7461A()
'GPIB to Offline

    sts = illoc(AD7461A)                'Set LocalMode
    sts = ilonl(AD7461A, 0)             'GPIB Offline

End Sub

Private Sub GetMeasure()
'Get Measure Voltage

    Call mSecSleep(ResetWaitTime)
    
    sts = ilrd(AD7461A, GetData, 25)    'Get Measure

End Sub

Public Sub ContinueMode()
'Start From Break Point

    If Worksheets(ReadSheetNameInfo).Cells(7, 4) <> 0 Then
        msgtext = "The last time ,You stopped in VoltCheckProccess" + vbCrLf + vbCrLf + _
                  "[Yes]    Start VoltCheck from last Condition" + vbCrLf + _
                  "             you lose Condition of the last time" + vbCrLf + _
                  "[No]     Start VoltCheck from your order" + vbCrLf + _
                  "[Cancel] Return to frm_PALS_VoltAdj_Main"
        ANS = MsgBox(msgtext, vbYesNoCancel + vbQuestion, "Inquire Message[VoltCheck]")
        
        If ANS = vbYes Then
            Flg_Break_voltcheck = True
        ElseIf ANS = vbNo Then
            Flg_Break_voltcheck = False
        ElseIf ANS = vbCancel Then
            Flg_End_voltcheck = True
        End If
    End If

End Sub

Public Sub Save_BreakPoint()
'Save Break Point

    Worksheets(ReadSheetNameInfo).Cells(7, 4).Value = Break_Counter
    
    Worksheets(ReadSheetNameInfo).Cells(8, 4).Value = frm_PALS_VoltAdj_Main.SitePrint.ListIndex
    Worksheets(ReadSheetNameInfo).Cells(9, 4).Value = frm_PALS_VoltAdj_Main.CONDITIONPrint.ListIndex
    Worksheets(ReadSheetNameInfo).Cells(10, 4).Value = frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex
    
    Worksheets(ReadSheetNameInfo).Cells(11, 4).Value = frm_PALS_VoltAdj_Main.SiteLockCheck.Value
    Worksheets(ReadSheetNameInfo).Cells(12, 4).Value = frm_PALS_VoltAdj_Main.ConditionLockCheck.Value
    Worksheets(ReadSheetNameInfo).Cells(13, 4).Value = frm_PALS_VoltAdj_Main.PinLockCheck.Value
    
    Worksheets(ReadSheetNameInfo).Cells(14, 4).Value = frm_PALS_VoltAdj_Main.ConfirmModeCheck.Value
    Worksheets(ReadSheetNameInfo).Cells(15, 4).Value = frm_PALS_VoltAdj_Main.DataLogCheck.Value
    Worksheets(ReadSheetNameInfo).Cells(16, 4).Value = frm_PALS_VoltAdj_Main.Onlyvoltagecheck.Value

End Sub

Public Sub Get_BreakPoint()
'Confirm For Break Point
    Dim y_cell

    Break_Counter = Worksheets(ReadSheetNameInfo).Cells(7, 4).Value
    Break_Counter_prev = Break_Counter

    frm_PALS_VoltAdj_Main.SitePrint.ListIndex = Worksheets(ReadSheetNameInfo).Cells(8, 4).Value
    frm_PALS_VoltAdj_Main.CONDITIONPrint.ListIndex = Worksheets(ReadSheetNameInfo).Cells(9, 4).Value
    frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex = Worksheets(ReadSheetNameInfo).Cells(10, 4).Value
    
    frm_PALS_VoltAdj_Main.SiteLockCheck.Value = Worksheets(ReadSheetNameInfo).Cells(11, 4).Value
    frm_PALS_VoltAdj_Main.ConditionLockCheck.Value = Worksheets(ReadSheetNameInfo).Cells(12, 4).Value
    frm_PALS_VoltAdj_Main.PinLockCheck.Value = Worksheets(ReadSheetNameInfo).Cells(13, 4).Value
    
    frm_PALS_VoltAdj_Main.ConfirmModeCheck.Value = Worksheets(ReadSheetNameInfo).Cells(14, 4).Value
    frm_PALS_VoltAdj_Main.DataLogCheck.Value = Worksheets(ReadSheetNameInfo).Cells(15, 4).Value
    frm_PALS_VoltAdj_Main.Onlyvoltagecheck.Value = Worksheets(ReadSheetNameInfo).Cells(16, 4).Value

    'BreakPoint Clear =======================
    For y_cell = 7 To 15
        Worksheets(ReadSheetNameInfo).Cells(y_cell, 4).Value = 0
    Next
    '========================================

End Sub
Public Sub Save_KeepPoint()
'Save Keep Point

    Worksheets(ReadSheetNameInfo).Cells(17, 4).Value = frm_PALS_VoltAdj_Main.SitePrint.ListIndex
'    Worksheets(ReadSheetNameInfo).Cells(18, 4).value = frm_PALS_VoltAdj_Main.CONDITIONPrint.ListIndex
    Worksheets(ReadSheetNameInfo).Cells(19, 4).Value = frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex
    
    Worksheets(ReadSheetNameInfo).Cells(20, 4).Value = frm_PALS_VoltAdj_Main.SiteLockCheck.Value
    Worksheets(ReadSheetNameInfo).Cells(21, 4).Value = frm_PALS_VoltAdj_Main.ConditionLockCheck.Value
    Worksheets(ReadSheetNameInfo).Cells(22, 4).Value = frm_PALS_VoltAdj_Main.PinLockCheck.Value
    
    Worksheets(ReadSheetNameInfo).Cells(23, 4).Value = frm_PALS_VoltAdj_Main.ConfirmModeCheck.Value
    Worksheets(ReadSheetNameInfo).Cells(24, 4).Value = frm_PALS_VoltAdj_Main.DataLogCheck.Value
    Worksheets(ReadSheetNameInfo).Cells(25, 4).Value = frm_PALS_VoltAdj_Main.Onlyvoltagecheck.Value

End Sub


Public Sub Get_KeepPoint()
'Confirm For Break Point

    frm_PALS_VoltAdj_Main.SitePrint.ListIndex = Worksheets(ReadSheetNameInfo).Cells(17, 4).Value
'    frm_PALS_VoltAdj_Main.CONDITIONPrint.ListIndex = Worksheets(ReadSheetNameInfo).Cells(18, 4).value
    frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex = Worksheets(ReadSheetNameInfo).Cells(19, 4).Value
    
    frm_PALS_VoltAdj_Main.SiteLockCheck.Value = Worksheets(ReadSheetNameInfo).Cells(20, 4).Value
    frm_PALS_VoltAdj_Main.ConditionLockCheck.Value = Worksheets(ReadSheetNameInfo).Cells(21, 4).Value
    frm_PALS_VoltAdj_Main.PinLockCheck.Value = Worksheets(ReadSheetNameInfo).Cells(22, 4).Value
    
    frm_PALS_VoltAdj_Main.ConfirmModeCheck.Value = Worksheets(ReadSheetNameInfo).Cells(23, 4).Value
    frm_PALS_VoltAdj_Main.DataLogCheck.Value = Worksheets(ReadSheetNameInfo).Cells(24, 4).Value
    frm_PALS_VoltAdj_Main.Onlyvoltagecheck.Value = Worksheets(ReadSheetNameInfo).Cells(25, 4).Value

End Sub

Public Sub Check_Flg()
'Check Flag Status

    Flg_SiteLock = frm_PALS_VoltAdj_Main.SiteLockCheck.Value
    Flg_ConditionLock = frm_PALS_VoltAdj_Main.ConditionLockCheck.Value
    Flg_PinLock = frm_PALS_VoltAdj_Main.PinLockCheck.Value
    Flg_ConfirmMode = frm_PALS_VoltAdj_Main.ConfirmModeCheck.Value
    Flg_Onlyvoltagecheck = frm_PALS_VoltAdj_Main.Onlyvoltagecheck.Value
    Flg_DataLog = frm_PALS_VoltAdj_Main.DataLogCheck.Value

End Sub

Public Sub SelectPin_check()
'Check before start

    'On Non checkRun ==================
    If Flg_PinLock = -1 And frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex <> -1 Then
        If CheckRun(frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex) = "OFF" Then
            msgtext = "Pin[CheckRun=OFF] selected..." + vbCrLf + _
                      "Check [CheckRun] or Change Pin!"
            ANS = MsgBox(msgtext & vbCrLf & "ErrCode.1-2-03-5-04", vbOKOnly + vbCritical, "Warning Message[VoltCheck]")
            
            frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex = -1
            frm_PALS_VoltAdj_Main.PinNumPrint.Value = ""
            DoEvents
            Flg_End_voltcheck = True
            Exit Sub
        End If
    End If
    '==================================

    'Check On Non Slect Pin when Lock =
    If (Flg_SiteLock = -1 And frm_PALS_VoltAdj_Main.SitePrint.ListIndex = -1) Or _
        (Flg_ConditionLock = -1 And frm_PALS_VoltAdj_Main.CONDITIONPrint.ListIndex = -1) Or _
        (Flg_PinLock = -1 And frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex = -1) Then
            msgtext = "Please Select Pin!![Pin Locked...] "
            ANS = MsgBox(msgtext & vbCrLf & "ErrCode.1-2-03-5-05", vbOKOnly + vbCritical, "Warning Message[VoltCheck]")
            Flg_End_voltcheck = True
            Exit Sub
    End If
    '==================================
    
    'Check On Slect Pin when Non Lock =
    If (Flg_SiteLock = 0 And frm_PALS_VoltAdj_Main.SitePrint.ListIndex <> -1) Or _
        (Flg_ConditionLock = 0 And frm_PALS_VoltAdj_Main.CONDITIONPrint.ListIndex <> -1) Or _
        (Flg_PinLock = 0 And frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex <> -1) Then
            msgtext = "You selected  Site/Condition/Pin" + vbCrLf + _
                      "But You Not Lock" + vbCrLf + vbCrLf + _
                      "You want to Check ALL ??"
            ANS = MsgBox(msgtext & vbCrLf & "ErrCode.1-2-03-5-06", vbOKCancel + vbQuestion, "Inquire Message[VoltCheck]")
            If ANS = vbCancel Then
                Flg_End_voltcheck = True
            End If
    End If
    '==================================
    
End Sub

Public Sub Print_Info()
'Printing Status On Display

    Dim Flg_tester As Integer
    Dim i_Tester As Integer
    Dim i_Condition As Integer
    Dim i_Pin As Integer
    
    Flg_End_voltcheck = False
    For i_Tester = 0 To Tester_exist_cnt - 1
        If Tester_Name(i_Tester) = Sw_Node Then
            Flg_tester = True
        End If
    Next
    If Flg_tester = False Then
        Call sub_errPALS("VoltAdjust Tool Run error" & "Not found Machine information in WorkSheet [ " & ReadSheetName & " ]" + vbCrLf + _
                  "Please Check WorkSheet" & " at 'Print_Info'", "1-2-04-5-07")
    End If

    'Printing TESTER
    frm_PALS_VoltAdj_Main.MachinePrint = Sw_Node

    'Printing SITE
    frm_PALS_VoltAdj_Main.SitePrint.Clear
    Dim site_loop As Integer
    For site_loop = 0 To nSite
        frm_PALS_VoltAdj_Main.SitePrint.AddItem site_loop
    Next site_loop
    frm_PALS_VoltAdj_Main.SitePrint.ListIndex = -1
    
    'Printing CONDITION
    frm_PALS_VoltAdj_Main.CONDITIONPrint.Clear
    For i_Condition = 0 To Condition_Number - 1
        frm_PALS_VoltAdj_Main.CONDITIONPrint.AddItem Test_Condition(i_Condition)
    Next i_Condition

    frm_PALS_VoltAdj_Main.CONDITIONPrint.ListIndex = -1
    If Now_Mode <> "" Then
        frm_PALS_VoltAdj_Main.ConditionLockCheck.Value = True
        frm_PALS_VoltAdj_Main.CONDITIONPrint.Value = Now_Mode
    End If

    'Printing PINNAME
    frm_PALS_VoltAdj_Main.PinNamePrint.Clear
    For i_Pin = 0 To pin_number - 1
        frm_PALS_VoltAdj_Main.PinNamePrint.AddItem PowerPinName(i_Pin)
    Next i_Pin
    frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex = -1
    
    frm_PALS_VoltAdj_Main.OffsetVolLimPrint.Text = LimitOffset
    frm_PALS_VoltAdj_Main.DiffVolLimPrint.Text = AdjustOffset

    DoEvents

End Sub

Public Sub Get_SelectInfo()
'Get Choice status(For Manual)

    Now_Site = frm_PALS_VoltAdj_Main.SitePrint.ListIndex
    Now_Condition = frm_PALS_VoltAdj_Main.CONDITIONPrint.ListIndex
    Now_PinName = frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex

End Sub

Public Function Check_Pin()
'Check connect Pin

    Dim Check_InputVolt(30) As Double
    Dim Check_OutputVolt(30) As Double
    Dim i_checkcnt
    CheckPin_flag(Now_Condition, Now_PinName, Now_Site) = 0

Retry:
    Flg_MeasSkip = False
    SetupVolt = Force_Voltage_vc(Now_Condition, Now_PinName)
    frm_PALS_VoltAdj_Main.SetupVolPrint = SetupVolt
    OffsetVolt = Offset_Voltage_vc(Now_Condition, Now_PinName, Now_Site)
    If OffsetVolt > 1 Then OffsetVolt = 0
    frm_PALS_VoltAdj_Main.OffsetVolPrint = OffsetVolt
    frm_PALS_VoltAdj_Main.DiffVolPrint = ""
    frm_PALS_VoltAdj_Main.InputVolPrint = ""
    frm_PALS_VoltAdj_Main.MeasVolPrint = ""
    
    frm_PALS_VoltAdj_Main.Message = "Wait For Start Pin Check"
    DoEvents
    
    msgtext = "Wait For Start Pin Check..." & vbCrLf & vbCrLf & _
              "Please Connect to TP[ " & PowerPinNum(Now_Site, Now_PinName) & " ]" & vbCrLf & _
              "Are You Ready??"
    ANS = MsgBox(msgtext, vbOKCancel + vbInformation, "Navigation Massage[VoltCheck]")
    
    If ANS = vbCancel Then
        CheckPin_flag(Now_Condition, Now_PinName, Now_Site) = 1
        Flg_Stop_voltcheck = True
        Exit Function
    End If
    
    If Flg_Onlyvoltagecheck = -1 Then Exit Function
    
    frm_PALS_VoltAdj_Main.Message = "Now in Proccess[PinCheck]....."
    DoEvents
    
    'Processing For Check connect Pin
    For i_checkcnt = 1 To 2
        Check_InputVolt(i_checkcnt) = SetupVolt + OffsetVolt + 0.1 * (i_checkcnt - 2) '100mV swing
        
        'Imposing Voltage
        If Flag_debug_Voltage <> 1 Then Call impose_volt(Check_InputVolt(i_checkcnt), Now_PinName, Now_Condition, Now_Site, 0)
        
        'Read connect Pin Voltage
        If Flag_debug_Voltage <> 1 Then Call GetMeasure
        If Flag_debug_Voltage <> 1 Then Call GetMeasure
        
        If Left(GetData, 5) = "DCV_ " Then GetData = Mid(GetData, 6, Len(GetData) - 5)
        Check_OutputVolt(i_checkcnt) = val(Mid(GetData, 1, 13))
        
        If Flag_debug_Voltage = 1 Then
            Check_OutputVolt(1) = 5#
            Check_OutputVolt(2) = 5.1
        End If
        
        If Flag_debug_Voltage = 1 Then Debug.Print "Checkpin-" & i_checkcnt & " " & Check_OutputVolt(i_checkcnt)
        If Flg_DataLog = -1 Then Print #FileNo_vc, "Checkpin-" & CStr(i_checkcnt) & " " & CStr(Check_OutputVolt(i_checkcnt))
    Next i_checkcnt

    'Check to Change Voltage amount
    OffsetVolt = (Check_OutputVolt(2) - Check_OutputVolt(1)) / 0.1

    'Judge to Change Voltage amount
    If (OffsetVolt < 0.9) Or (1.1 < OffsetVolt) Then
        CheckPin_flag(Now_Condition, Now_PinName, Now_Site) = 1
        If Flg_DataLog = -1 Then Print #FileNo_vc, "!!!! PinCheck NG !!!!"
        msgtext = "Pin Check ERROR!" + vbCrLf + vbCrLf + _
                  "Please Check Pin connection" + vbCrLf + _
                  "[Cancel] Stop Check this Pin"
        ANS = MsgBox(msgtext & vbCrLf & "ErrCode.1-2-05-3-08", vbRetryCancel + vbCritical, "Warning Message[VoltCheck]")
        If ANS = vbRetry Then
            GoTo Retry
        End If
        If ANS = vbCancel Then
            Flg_MeasSkip = True
            Exit Function
        End If
    Else
        CheckPin_flag(Now_Condition, Now_PinName, Now_Site) = 0
        If Flg_DataLog = -1 Then Print #FileNo_vc, "---- PinCheck OK ----"
    End If
    

    'Message On matching connect Pin
    frm_PALS_VoltAdj_Main.Message = "Pin Check OK" + vbCrLf + vbCrLf + _
                            "  Next...Check Offset"
    DoEvents

    'Viewable On Flow
    frm_PALS_VoltAdj_Main.ConnectCheck.Visible = True

End Function

Public Function Adjust_Volt()
'Adjust Voltage

    Dim Adjust_Count As Integer

    'Printing Setting Voltage
    If Flg_DataLog = -1 Then Print #FileNo_vc, "SetupVolt  " & CStr(SetupVolt)
    If Flag_debug_Voltage = 1 Then Debug.Print "SetupVolt  " & SetupVolt
    frm_PALS_VoltAdj_Main.Message = "Now Voltage Checking...."
    DoEvents
    
    Adjust_Count = 0
    
    OffsetVolt = Offset_Voltage_vc(Now_Condition, Now_PinName, Now_Site)
    If OffsetVolt > 1 Then OffsetVolt = 0
    frm_PALS_VoltAdj_Main.OffsetVolPrint = OffsetVolt
    DiffVolt = 0
    Do
        If Flg_DataLog = -1 Then Print #FileNo_vc, ""
        
        Adjust_Count = Adjust_Count + 1
                
        'Voltage Clear
        frm_PALS_VoltAdj_Main.MeasVolPrint.Value = ""
        frm_PALS_VoltAdj_Main.DiffVolPrint.Value = ""
        frm_PALS_VoltAdj_Main.Applycheck.Visible = False

        'Reflect Offset
        OffsetVolt = OffsetVolt + DiffVolt
        frm_PALS_VoltAdj_Main.OffsetVolPrint = OffsetVolt
        
        InputVolt = SetupVolt + OffsetVolt
        frm_PALS_VoltAdj_Main.InputVolPrint = InputVolt

        If Flg_DataLog = -1 Then Print #FileNo_vc, "OffsetVolt " & CStr(OffsetVolt)
        If Flg_DataLog = -1 Then Print #FileNo_vc, "Input      " & CStr(InputVolt)
        If Flag_debug_Voltage = 1 Then Debug.Print "OffsetVolt " & OffsetVolt
        If Flag_debug_Voltage = 1 Then Debug.Print "Input      " & InputVolt
        DoEvents

'        Debug.Print InputVolt
        'Imposing Voltage
        If Flag_debug_Voltage <> 1 And Flg_Onlyvoltagecheck <> -1 Then Call impose_volt(InputVolt, Now_PinName, Now_Condition, Now_Site, 0)
        
        'Viewable On Flow
        frm_PALS_VoltAdj_Main.Applycheck.Visible = True
        DoEvents
        
        'Read connect Pin Voltage
        If Flag_debug_Voltage <> 1 Then
            Call GetMeasure
            Call GetMeasure
            If Left(GetData, 5) = "DCV_ " Then GetData = Mid(GetData, 6, Len(GetData) - 5)
            MeasVolt = val(Mid(GetData, 1, 13))
            If MeasVolt <> 0 Then MeasVolt = Format$(MeasVolt, "#.####")
        End If
'        Debug.Print MeasVolt
        If Flag_debug_Voltage = 1 Then Call mSecSleep(500)

        If Flag_debug_Voltage = 1 Then
            If Adjust_Count = 1 Then MeasVolt = SetupVolt - 0.01
            If Adjust_Count = 2 Then MeasVolt = SetupVolt + 0.0001
        End If
        
        'Printing Measure Voltage
        frm_PALS_VoltAdj_Main.MeasVolPrint = MeasVolt
        DoEvents

        'Output Datalog to Textfile
        If Flag_debug_Voltage = 1 Then Debug.Print "MeasVolt   " & MeasVolt
        If Flg_DataLog = -1 Then Print #FileNo_vc, "MeasVolt   " & CStr(MeasVolt)
        
        'Computing Diff Amount
        If SetupVolt - MeasVolt <> 0 Then
            DiffVolt = Format$(SetupVolt - MeasVolt, "#.####")    'Range 0.1mV
        Else
            DiffVolt = 0#
        End If
        
        frm_PALS_VoltAdj_Main.DiffVolPrint = DiffVolt
        If Flag_debug_Voltage = 1 Then Debug.Print "DiffVolt   " & DiffVolt
        If Flg_DataLog = -1 Then Print #FileNo_vc, "DiffVolt   " & CStr(DiffVolt)
        DoEvents

        'On Not Adjust
        If Flg_ConfirmMode = -1 Or Flg_Onlyvoltagecheck = -1 Then
            OffsetVolt = OffsetVolt + DiffVolt
            GoTo End_Adjust
        End If

        'Judge to Offset amount
        If Abs(OffsetVolt + DiffVolt) > LimitOffset Then
            msgtext = "Offset over limit [" & LimitOffset & "]V!!" & vbCrLf & vbCrLf & _
                      "[Retry]  Retry adjust this Pin" & vbCrLf & _
                      "[Cancel] Stop adjust this Pin & Goto NextPin"
            ANS = MsgBox(msgtext & vbCrLf & "ErrCode.1-2-06-3-09", vbRetryCancel + vbCritical, "Warning Message[VoltCheck]")
            If ANS = vbCancel Then
                OffsetVolt = OffsetVolt + DiffVolt
                If Flg_DataLog = -1 Then Print #FileNo_vc, "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
                If Flg_DataLog = -1 Then Print #FileNo_vc, "!!!!    Check NG(LimitOver)    !!!!"
                If Flg_DataLog = -1 Then Print #FileNo_vc, "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
                GoTo End_Adjust
            End If
        End If

        'Checking Loop number
        If Adjust_Count >= 10 Then
            msgtext = "Over 10 Loop for OffsetCheck " + vbCrLf + _
                      "Finish this check , and ..Do Next check..."
            ANS = MsgBox(msgtext & vbCrLf & "ErrCode.1-2-06-3-10", vbOKOnly + vbCritical, "Warning Message[VoltCheck]")
            If Flg_DataLog = -1 Then Print #FileNo_vc, "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
            If Flg_DataLog = -1 Then Print #FileNo_vc, "!!!!    Check NG(LoopOver)    !!!!"
            If Flg_DataLog = -1 Then Print #FileNo_vc, "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
            GoTo End_Adjust
        End If

    Loop While Abs(DiffVolt) >= AdjustOffset
    
    If Flg_DataLog = -1 Then Print #FileNo_vc, "------------------------"
    If Flg_DataLog = -1 Then Print #FileNo_vc, "----    Check OK    ----"
    If Flg_DataLog = -1 Then Print #FileNo_vc, "------------------------"

End_Adjust:
    If Flg_DataLog = -1 Then Print #FileNo_vc, ""

    'Reflect To Worksheet Offset
    Offset_Voltage_vc(Now_Condition, Now_PinName, Now_Site) = OffsetVolt

    'Viewable On Flow
    frm_PALS_VoltAdj_Main.Message = "Voltage Check Finished...."
    If Flg_ConfirmMode <> -1 And Flg_Onlyvoltagecheck <> -1 Then frm_PALS_VoltAdj_Main.OffsetCheck.Visible = True

    DoEvents
    Call mSecSleep(100)

End Function

Public Function impose_volt(voltage As Double, Pin As Integer, Condition As Integer, mySite As Long, Allsite As Integer)
'   Allsite = 1 ... Run All site
'   Allsite = 0 ... Run One site

    Dim ClampCurrentRange As String
    Dim ForceVoltageRange As String
    Dim Channels() As Long              'For PPMU
    
    Dim ClampCurrentRangePPMU As Long
    Dim numChans As Long
    Dim strError As String
    Dim numSites As Long

    Dim flg_impose_check As Boolean
    flg_impose_check = False
    
    '================================================================
    'For APMU Imporse_volt Start ====================================
    #If APMU_USE <> 0 Then
    If Pin_Resource(Pin) = "APMU" Then
    
        'ClampCurrent Range Setup ==========
        If Abs(Clamp_Current_vc(Condition, Pin)) <= 40 * uA Then
            ClampCurrentRange = apmu40uA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 200 * uA Then
            ClampCurrentRange = apmu200uA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 1 * mA Then
            ClampCurrentRange = apmu1mA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 5 * mA Then
            ClampCurrentRange = apmu5mA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 50 * mA Then
            ClampCurrentRange = apmu50mA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 400 * mA Then   'For Gang
            ClampCurrentRange = apmu50mA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) > 400 * mA Then
            MsgBox "ClampCurrent Range Over" & vbCrLf & "ErrCode.1-2-07-5-11", vbOKOnly + vbCritical, "ClampCurrent Range Over[VoltCheck]"
        End If
        '===================================
        
        'ForceVoltage Range Setup ==========
        If Abs(voltage) <= 2 * V Then
            ForceVoltageRange = apmu2V
        ElseIf Abs(voltage) <= 5 * V Then
            ForceVoltageRange = apmu5V
        ElseIf Abs(voltage) <= 10 * V Then
            ForceVoltageRange = apmu10V
        ElseIf Abs(voltage) <= 35 * V Then
            ForceVoltageRange = apmu35V
        ElseIf Abs(voltage) > 35 * V Then
            MsgBox "ForceVoltage Range Over" & vbCrLf & "ErrCode.1-2-07-5-12", vbOKOnly + vbCritical, "ForceVoltage Range Over[VoltCheck]"
        End If
        '===================================
        
        If Allsite = 1 Then
            With TheHdw.APMU.Pins(PowerPinName(Pin))
                .alarm = False
                .ModeFVMI (ClampCurrentRange)
                .ClampCurrent(ClampCurrentRange) = Clamp_Current_vc(Condition, Pin)
                .ForceVoltage(ForceVoltageRange) = voltage
                .relay = True
                .Gate = True
            End With
        ElseIf Allsite = 0 Then
            With TheHdw.APMU.Chans(Pin_Chans(mySite, Pin))
                .alarm = False
                .ModeFVMI (ClampCurrentRange)
                .ClampCurrent(ClampCurrentRange) = Clamp_Current_vc(Condition, Pin)
                .ForceVoltage(ForceVoltageRange) = voltage
                .relay = True
                .Gate = True
            End With
        End If
        flg_impose_check = True
    End If
    #End If
    'For APMU Imporse_volt End ======================================
    '================================================================



    '================================================================
    'For BPMU Imporse_volt Start ====================================
    #If BPMU_USE <> 0 Then
    Dim ClampCurrentRangeBPMU As BpmuIRange
    Dim ForceVoltageRangeBPMU As BpmuVRange
    If Pin_Resource(Pin) = "BPMU" Then

        'ClampCurrent Range Setup ==========
        If Abs(Clamp_Current_vc(Condition, Pin)) <= 2 * uA Then
            ClampCurrentRangeBPMU = bpmu2uA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 20 * uA Then
            ClampCurrentRangeBPMU = bpmu20uA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 200 * uA Then
            ClampCurrentRangeBPMU = bpmu200uA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 2 * mA Then
            ClampCurrentRangeBPMU = bpmu2mA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 20 * mA Then
            ClampCurrentRangeBPMU = bpmu20mA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) > 20 * mA Then
            MsgBox "ClampCurrent Range Over" & vbCrLf & "ErrCode.1-2-07-5-13", vbOKOnly + vbCritical, "ClampCurrent Range Over[VoltCheck]"
        End If
        '===================================

        'ForceVoltage Range Setup ==========
        If Abs(voltage) <= 2 * V Then
            ForceVoltageRangeBPMU = bpmu2V
        ElseIf Abs(voltage) <= 5 * V Then
            ForceVoltageRangeBPMU = bpmu5V
        ElseIf Abs(voltage) <= 10 * V Then
            ForceVoltageRangeBPMU = bpmu10V
        ElseIf Abs(voltage) <= 24 * V Then
            ForceVoltageRangeBPMU = bpmu24V
        ElseIf Abs(voltage) > 24 * V Then
            MsgBox "ForceVoltage Range Over" & vbCrLf & "ErrCode.1-2-07-5-14", vbOKOnly + vbCritical, "ForceVoltage Range Over[VoltCheck]"
        End If
        '===================================

        If Allsite = 1 Then
            With TheHdw.BPMU.Pins(PowerPinName(Pin))
                .ClampCurrent(ClampCurrentRangeBPMU) = Clamp_Current_vc(Condition, Pin)
                .ForceVoltage(ForceVoltageRangeBPMU) = voltage
                Call .ModeFVMI(ClampCurrentRangeBPMU, ForceVoltageRangeBPMU)
            End With
        ElseIf Allsite = 0 Then
            With TheHdw.BPMU.Chans(Pin_Chans(mySite, Pin))
                .ClampCurrent(ClampCurrentRangeBPMU) = Clamp_Current_vc(Condition, Pin)
                .ForceVoltage(ForceVoltageRangeBPMU) = voltage
                Call .ModeFVMI(ClampCurrentRangeBPMU, ForceVoltageRangeBPMU)
            End With
        End If
        flg_impose_check = True
    End If
    #End If
    'For BPMU Imporse_volt End =======================================
    '================================================================



    '================================================================
    'For PPMU Imporse_volt Start ====================================
    #If PPMU_USE <> 0 Then
    If Pin_Resource(Pin) = "PPMU" Then

        'ClampCurrent Range Setup ==========
        If Abs(Clamp_Current_vc(Condition, Pin)) <= 2 * uA Then
            ClampCurrentRangePPMU = ppmu2uA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 20 * uA Then
            ClampCurrentRangePPMU = ppmu20uA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 200 * uA Then
            ClampCurrentRangePPMU = ppmu200uA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 2 * mA Then
            ClampCurrentRangePPMU = ppmu2mA
        '>>>2013/3/4 M.IMAMURA Changed
        '####For IP750EX
        #If HSD200_USE <> 0 Then
            ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 50 * mA Then
                ClampCurrentRangePPMU = ppmu50mA
            ElseIf Abs(Clamp_Current_vc(Condition, Pin)) > 50 * mA Then
        #Else
            '####For OtherSystem
            ElseIf Abs(Clamp_Current_vc(Condition, Pin)) > 2 * mA Then
        #End If
        '<<<2013/3/4 M.IMAMURA Changed
            MsgBox "ClampCurrent Range Over" & vbCrLf & "ErrCode.1-2-07-5-15", vbOKOnly + vbCritical, "ClampCurrent Range Over[VoltCheck]"
        End If
        '===================================

        'ForceVoltage Range Setup ==========
        If (voltage <= -2 * V) Or (voltage >= 7 * V) Then
            MsgBox "ForceVoltage Range Over" & vbCrLf & "ErrCode.1-2-07-5-16", vbOKOnly + vbCritical, "ForceVoltage Range Over[VoltCheck]"
        End If
        '===================================

        If Allsite = 1 Then
            With TheHdw.PPMU.Pins(PowerPinName(Pin))
                .ForceVoltage(ClampCurrentRangePPMU) = voltage
                .Connect
            End With
        ElseIf Allsite = 0 Then
            Call TheExec.DataManager.GetChanList(PowerPinName(Pin), mySite, chIO, Channels, numChans, numSites, strError)     'Get Channel#
            '>>>2011/8/22 M.IMAMURA Changed
            'With TheHdw.PPMU.chans(Channels(mysite))
            With TheHdw.PPMU.Chans(Channels(0))
            '<<<2011/8/22 M.IMAMURA Changed
                .ForceVoltage(ClampCurrentRangePPMU) = voltage
                .Connect
            End With
        End If
        flg_impose_check = True
    End If
    #End If
    'For PPMU Imporse_volt End ======================================
    '================================================================



    '================================================================
    'For HDVIS Imporse_volt Start ===================================
    #If HDVIS_USE <> 0 Then
    If Pin_Resource(Pin) = "HDVIS" Then

        'ClampCurrent Range Setup ==========
        If Abs(Clamp_Current_vc(Condition, Pin)) <= 5 * uA Then
            ClampCurrentRange = hdvis5uA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 50 * uA Then
            ClampCurrentRange = hdvis50uA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 500 * uA Then
            ClampCurrentRange = hdvis500uA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 5 * mA Then
            ClampCurrentRange = hdvis5mA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 50 * mA Then
            ClampCurrentRange = hdvis50mA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 800 * mA Then   'For Gang
            ClampCurrentRange = hdvis200mA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) > 800 * mA Then
            MsgBox "ClampCurrent Range Over" & vbCrLf & "ErrCode.1-2-07-5-17", vbOKOnly + vbCritical, "ClampCurrent Range Over[VoltCheck]"
        End If
        '===================================

        If Allsite = 1 Then
            With TheHdw.HDVIS.Pins(PowerPinName(Pin))
                .alarm(hdvisAlarmOpenDGS) = False
                .alarm(hdvisAlarmOverLoad) = False
                .ModeFVMI (ClampCurrentRange)
                .ClampCurrent(ClampCurrentRange) = Clamp_Current_vc(Condition, Pin)
                .ForceVoltage(HDVIS10V) = voltage
                .relay = True
                .Gate = True
            End With
        ElseIf Allsite = 0 Then
            With TheHdw.HDVIS.Chans(Pin_Chans(mySite, Pin))
                .alarm(hdvisAlarmOpenDGS) = False
                .alarm(hdvisAlarmOverLoad) = False
                .ModeFVMI (ClampCurrentRange)
                .ClampCurrent(ClampCurrentRange) = Clamp_Current_vc(Condition, Pin)
                .ForceVoltage(HDVIS10V) = voltage
                .relay = True
                .Gate = True
            End With
        End If
        flg_impose_check = True
    End If
    #End If
    'For HDVIS Imporse_volt End =====================================
    '================================================================


    '================================================================
    'For DPS Imporse_volt Start =====================================
    #If DPS_USE <> 0 Then
    If Pin_Resource(Pin) = "DPS" Then

        'ClampCurrent Range Setup ==========
        If Abs(Clamp_Current_vc(Condition, Pin)) <= 50 * uA Then
            ClampCurrentRange = dps50uA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 500 * uA Then
            ClampCurrentRange = dps500ua
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 10 * mA Then
            ClampCurrentRange = dps10ma
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 100 * mA Then
            ClampCurrentRange = dps100mA
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 1 * A Then
            ClampCurrentRange = dps1a
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) <= 4 * A Then   'For Gang
            ClampCurrentRange = dps1a
        ElseIf Abs(Clamp_Current_vc(Condition, Pin)) > 4 * A Then
            MsgBox "ClampCurrent Range Over" & vbCrLf & "ErrCode.1-2-07-5-18", vbOKOnly + vbCritical, "ClampCurrent Range Over[VoltCheck]"
        End If
        '===================================

        If Allsite = 1 Then
            With TheHdw.DPS.Pins(PowerPinName(Pin))
                .CurrentRange = ClampCurrentRange
                .CurrentLimit = Clamp_Current_vc(Condition, Pin)
                .forceValue(dpsPrimaryVoltage) = voltage
            End With
        ElseIf Allsite = 0 Then
            With TheHdw.DPS.Chans(Pin_Chans(mySite, Pin))
                .CurrentRange = ClampCurrentRange
                .CurrentLimit = Clamp_Current_vc(Condition, Pin)
                .forceValue(dpsPrimaryVoltage) = voltage
            End With
        End If
        flg_impose_check = True
    End If
    #End If
    'For DPS Imporse_volt End =======================================
    '================================================================

    'Impose Check
    If flg_impose_check = False Then
        msgtext = "ImposeVolt ERROR!" + vbCrLf + vbCrLf + _
              "Please Check Job Setting!!" + vbCrLf + _
              " Sheet[" & PinSheetname & "]" + vbCrLf + _
              " Sheet[" & ReadSheetName & "]" + vbCrLf + _
              " Function[impose_volt]"

        ANS = MsgBox(msgtext & vbCrLf & "ErrCode.1-2-07-4-19", vbOKOnly + vbCritical, "Warning Message[VoltCheck]")
    End If
End Function

Public Function Get_Power_Condition()

    Dim Start_cell_x As Long    'Read start Cell X adress
    Dim Start_cell_y As Long    'Read start Cell Y adress
    Dim Cha_cell_count_x As Integer
    Dim Cha_cell_count_y As Integer
    Dim Cell_count_x As Long    'Count X Cell
    Dim Cell_count_y As Long    'Count Y Cell
    Dim Counter_Cond    As Integer
    Dim Condition_num   As Integer
    Dim Counter_Pin     As Integer
    Dim Counter_Site    As Integer
    Dim Counter_Node    As Integer
    
    Set Cnum = Nothing
    Set Cnum = New Collection
    Set PNum = Nothing
    Set PNum = New Collection
    
    '>>> 2011/6/30 M.Imamura
    If sub_SheetNameCheck(PinSheetnameChans) = True Then
        PinSheetname = PinSheetnameChans
    Else
        PinSheetname = PinSheetnameChannel
    End If
    '<<< 2011/5/30 M.Imamura
    
    'Get Status ===========================================
    Counter_Cond = 0
    Start_cell_x = Start_x
    Start_cell_y = Start_y
    Cell_count_x = 0
    Cell_count_y = 0
    If Worksheets(ReadSheetName).Cells(Start_cell_y, Start_cell_x).Value = "" Then Exit Function
    
    Do Until Worksheets(ReadSheetName).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x).Value = ""
    
        ReDim Preserve Test_Condition(Counter_Cond) 'Dynamic Array
        
        Test_Condition(Counter_Cond) = Worksheets(ReadSheetName).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x).Value
        Cell_count_y = Cell_count_y + 1
        
        If Test_Condition(Counter_Cond) <> Worksheets(ReadSheetName).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x).Value Then
            Cnum.Add Item:=CStr(Counter_Cond), key:=Test_Condition(Counter_Cond)
            Counter_Cond = Counter_Cond + 1
        End If
        
    Loop
    Condition_Number = Counter_Cond  'Condition Total Number
    '======================================================
    
    'Get Tester# and Site Total Number ====================
    Counter_Node = 0
    Counter_Site = 0
    Start_cell_x = Start_x + 2
    Start_cell_y = Start_y + 2
    Cell_count_x = 0
    Cell_count_y = 0
    Tester_exist_cnt = 0
    
    ReDim Preserve Next_Tester_YAddress(Counter_Node)
    
    Next_Tester_YAddress(Counter_Node) = Start_cell_y
    
    Do Until Worksheets(ReadSheetName).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x).Value = ""
        
        ReDim Preserve Tester_Name(Counter_Node) 'Dynamic Array
        ReDim Preserve Site_Number(Counter_Node) 'Dynamic Array

        Tester_Name(Counter_Node) = Worksheets(ReadSheetName).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x).Value
                    
            If Tester_Name(Counter_Node) = Worksheets(ReadSheetName).Cells(Start_cell_y + Cell_count_y + 1, Start_cell_x + Cell_count_x).Value Then
                Cell_count_y = Cell_count_y + 1
                Counter_Site = Counter_Site + 1
                Site_Number(Counter_Node) = Counter_Site
            Else
                Cell_count_y = Cell_count_y + 1
                Counter_Node = Counter_Node + 1
                Counter_Site = 0
                
                ReDim Preserve Next_Tester_YAddress(Counter_Node)
                Next_Tester_YAddress(Counter_Node) = Start_cell_y + Cell_count_y
                Tester_exist_cnt = Tester_exist_cnt + 1
            End If
    Loop
    Next_Condition_YAddress = Next_Tester_YAddress(Counter_Node) - Next_Tester_YAddress(0) + 2     'Condition Cell Range
    '======================================================

    'Get Pin Name Data ====================================
    Counter_Pin = 0
    Start_cell_x = 6
    Start_cell_y = 4
    Cell_count_x = 0
    Cell_count_y = 0
    Cha_cell_count_x = 0
    Cha_cell_count_y = 0
    
    Do Until Worksheets(ReadSheetName).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x).Value = ""
        
        ReDim Preserve PowerPinName(Counter_Pin)        'Dynamic Array
        ReDim Preserve Pin_Resource(Counter_Pin)        'Dynamic Array
        ReDim Preserve Pin_Chans(nSite, Counter_Pin)    'Dynamic Array
        
        PowerPinName(Counter_Pin) = Worksheets(ReadSheetName).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x).Value
        
        '########PinNameCheck
        Dim intLooppin1 As Integer
        Dim intLooppin2 As Integer
        
        For intLooppin1 = 0 To Counter_Pin
        For intLooppin2 = 0 To Counter_Pin
        If (intLooppin2 <> intLooppin1) And ("P_" & PowerPinName(intLooppin1) = PowerPinName(intLooppin2) Or PowerPinName(intLooppin1) = "P_" & PowerPinName(intLooppin2) Or PowerPinName(intLooppin1) = PowerPinName(intLooppin2)) Then
            msgtext = "Unexpected Pinname is existing!!" + vbCrLf + vbCrLf + _
                    "Pin[ " & PowerPinName(intLooppin1) & " ]" + vbCrLf + _
                    "Pin[ " & PowerPinName(intLooppin2) & " ]"
            ANS = MsgBox(msgtext & vbCrLf & "ErrCode.1-2-08-5-20", vbOKOnly + vbCritical, "Error Massage[VoltCheck]")
            Exit Function
        End If
        Next intLooppin2
        Next intLooppin1
        '#####################
        
        PNum.Add Item:=CStr(Counter_Pin), key:=PowerPinName(Counter_Pin)
        
        
        Do Until Worksheets(PinSheetname).Cells(Cha_start_cell_y + Cha_cell_count_y, Cha_start_cell_x + Cha_cell_count_x).Value = ""
                    
            If "P_" & PowerPinName(Counter_Pin) = Worksheets(PinSheetname).Cells(Cha_start_cell_y + Cha_cell_count_y, Cha_start_cell_x + Cha_cell_count_x).Value Then
                PNum.Remove Counter_Pin + 1
                PowerPinName(Counter_Pin) = "P_" & PowerPinName(Counter_Pin)
                PNum.Add Item:=CStr(Counter_Pin), key:=PowerPinName(Counter_Pin)
            End If
            
            If PowerPinName(Counter_Pin) = Worksheets(PinSheetname).Cells(Cha_start_cell_y + Cha_cell_count_y, Cha_start_cell_x + Cha_cell_count_x).Value Then
                Pin_Resource(Counter_Pin) = Worksheets(PinSheetname).Cells(Cha_start_cell_y + Cha_cell_count_y, Cha_start_cell_x + Cha_cell_count_x + 2).Value
                If Pin_Resource(Counter_Pin) = "I/O" Then Pin_Resource(Counter_Pin) = "PPMU"
                
                If Pin_Resource(Counter_Pin) = "HDVIS" Then
                    ' For HDVIS
                    Dim site_loop As Integer
                    For site_loop = 0 To nSite
                        Pin_Chans(site_loop, Counter_Pin) = val(Right(Worksheets(PinSheetname).Cells(Cha_start_cell_y + Cha_cell_count_y, Cha_start_cell_x + Cha_cell_count_x + 3 + site_loop).Value, Len(Worksheets(PinSheetname).Cells(Cha_start_cell_y + Cha_cell_count_y, Cha_start_cell_x + Cha_cell_count_x + 3 + site_loop).Value) - 5))
                    Next
                ElseIf Pin_Resource(Counter_Pin) = "PPMU" Then
                    ' For PPMU
                    For site_loop = 0 To nSite
                        Pin_Chans(site_loop, Counter_Pin) = val(Right(Worksheets(PinSheetname).Cells(Cha_start_cell_y + Cha_cell_count_y, Cha_start_cell_x + Cha_cell_count_x + 3 + site_loop).Value, Len(Worksheets(PinSheetname).Cells(Cha_start_cell_y + Cha_cell_count_y, Cha_start_cell_x + Cha_cell_count_x + 3 + site_loop).Value) - 2))
                    Next
                ElseIf Pin_Resource(Counter_Pin) = "DPS" Then
                    ' For DPS
                    For site_loop = 0 To nSite
                        Pin_Chans(site_loop, Counter_Pin) = val(Right(Worksheets(PinSheetname).Cells(Cha_start_cell_y + Cha_cell_count_y, Cha_start_cell_x + Cha_cell_count_x + 3 + site_loop).Value, Len(Worksheets(PinSheetname).Cells(Cha_start_cell_y + Cha_cell_count_y, Cha_start_cell_x + Cha_cell_count_x + 3 + site_loop).Value) - 3))
                    Next
                Else
                    ' For APMU/BPMU
                    For site_loop = 0 To nSite
                        Pin_Chans(site_loop, Counter_Pin) = val(Right(Worksheets(PinSheetname).Cells(Cha_start_cell_y + Cha_cell_count_y, Cha_start_cell_x + Cha_cell_count_x + 3 + site_loop).Value, 3))
                    Next
                End If
            End If
            Cha_cell_count_y = Cha_cell_count_y + 1
        Loop
        If Pin_Resource(Counter_Pin) = "" Then Pin_Resource(Counter_Pin) = "NonList"
            
        Cha_cell_count_y = 0
        Cell_count_x = Cell_count_x + 1
        Counter_Pin = Counter_Pin + 1 '
    Loop
    pin_number = Counter_Pin
    '======================================================

    'Get Pin# Data ========================================
    Start_cell_x = 6
    Start_cell_y = 6
    Cell_count_x = 0
    Cell_count_y = 0
    
    ReDim Preserve PowerPinNum(nSite, pin_number - 1) 'Dynamic Array
    
    For Cell_count_y = 0 To nSite
        For Cell_count_x = 0 To pin_number - 1
            PowerPinNum(Cell_count_y, Cell_count_x) = Worksheets(ReadSheetNameInfo).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x).Value
        Next
    Next
    '======================================================

    'Get CheckRun Data ====================================
    Counter_Pin = 0
    Start_cell_x = 6
    Start_cell_y = 5
    
    ReDim Preserve CheckRun(pin_number - 1) 'Dynamic Array
    
    For Counter_Pin = 0 To pin_number - 1
        CheckRun(Counter_Pin) = Worksheets(ReadSheetNameInfo).Cells(Start_cell_y, Start_cell_x + Counter_Pin).Value
    Next
    '======================================================

    'Get ForceVoltage and ClampCurrent ====================
    Counter_Pin = 0
    Counter_Cond = 0
    Start_cell_x = Start_x + 4
    Start_cell_y = Start_y
    Cell_count_x = 0
    Cell_count_y = 0

    ReDim Preserve Force_Voltage_vc(Condition_Number, pin_number - 1) 'Dynamic Array
    ReDim Preserve Clamp_Current_vc(Condition_Number, pin_number - 1) 'Dynamic Array

    Do Until Worksheets(ReadSheetName).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x - 3).Value = ""

        If Worksheets(ReadSheetName).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x - 3).Value = "Force Voltage[V]" Then

            Do Until Worksheets(ReadSheetName).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x).Value = ""
                Force_Voltage_vc(Counter_Cond, Counter_Pin) = Worksheets(ReadSheetName).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x).Value
                Clamp_Current_vc(Counter_Cond, Counter_Pin) = Worksheets(ReadSheetName).Cells(Start_cell_y + Cell_count_y + 1, Start_cell_x + Cell_count_x).Value

                Cell_count_x = Cell_count_x + 1
                Counter_Pin = Counter_Pin + 1 '
            Loop
            Cell_count_x = 0
            Counter_Pin = 0
            Counter_Cond = Counter_Cond + 1
        End If

        Cell_count_y = Cell_count_y + 1

    Loop
    '======================================================

    'Get Offset Data ======================================
    Start_cell_x = Start_x + 4
    Start_cell_y = Start_y + 2
    Cell_count_x = 0
    Cell_count_y = 0

    For Counter_Node = 0 To Tester_exist_cnt - 1
        If Tester_Name(Counter_Node) = Sw_Node Or Counter_Node = Tester_exist_cnt - 1 Then
            Start_cell_y = Next_Tester_YAddress(Counter_Node)
            ReDim Preserve Offset_Voltage_vc(Condition_Number, pin_number - 1, Site_Number(Counter_Node)) 'Dynamic Array
            ReDim Preserve CheckPin_flag(Condition_Number, pin_number - 1, Site_Number(Counter_Node)) 'Dynamic Array
            
            For Counter_Cond = 0 To Condition_Number - 1
                
                For Counter_Site = 0 To Site_Number(Counter_Node)
                    
                    For Counter_Pin = 0 To pin_number - 1
                        Offset_Voltage_vc(Counter_Cond, Counter_Pin, Counter_Site) = Worksheets(ReadSheetName).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x).Value
                        Cell_count_x = Cell_count_x + 1
                    Next Counter_Pin
                
                    Cell_count_x = 0
                    Cell_count_y = Cell_count_y + 1
                
                Next Counter_Site
                
                Cell_count_y = 0
                Start_cell_y = Start_cell_y + Next_Condition_YAddress
            
            Next Counter_Cond
        Exit For
        End If
    Next Counter_Node
    '======================================================
        
    Flg_voltsheet_read = 1

End Function

Public Sub offset_write()
'Writing offset in Worksheet

    frm_PALS_VoltAdj_Main.Message = "Offset Check Done...."
    DoEvents

'#########Not Adjust Mode
    If Flg_ConfirmMode = -1 Or Flg_Onlyvoltagecheck = -1 Then
'        Call Output_Offset(OutPutSheetname, 0)
        Call Output_Offset(OutPutSheetname, 1)
    
        msgtext = "VoltCheck Finished" + vbCrLf + vbCrLf + _
                  "Saved Offset to Sheet[ " & OutPutSheetname & " ]" + vbCrLf + _
                  "Please check Diiference"
        ANS = MsgBox(msgtext, vbOKOnly + vbInformation, "Information Massage[VoltCheck]")

        Worksheets(ReadSheetName).Select
    Else
'#########Do Adjust Mode
    '    Call Make_Sheet(ReadSheetName, OutPutSheetname)
        Call Backup_Offset(ReadSheetName, OutPutSheetname)
        Call Output_Offset(ReadSheetName, 1)
    
        msgtext = "VoltCheck Finished" + vbCrLf + vbCrLf + _
                  "Saved Offset to Sheet[ " & ReadSheetName & " ]" + vbCrLf + _
                  "Saved BackUp to Sheet[ " & OutPutSheetname & " ]"
        ANS = MsgBox(msgtext, vbOKOnly + vbInformation, "Information Massage[VoltCheck]")

        Worksheets(ReadSheetName).Select

    End If

End Sub

Public Function Make_Sheet(Original_Sheet As String, Copy_Sheet As String)
        
    Dim sheet_name As Variant
    Dim Sheet_Exist As Integer
    
    Sheet_Exist = False
    
    For Each sheet_name In Sheets
        If sheet_name.Name = Copy_Sheet Then
            Sheet_Exist = True
            Exit For
        End If
    Next
    
    If Sheet_Exist = True Then
        Worksheets(Copy_Sheet).Select
        Application.DisplayAlerts = False
        ActiveSheet.Delete
        Application.DisplayAlerts = True
    End If
        
    Worksheets(Original_Sheet).Copy After:=Worksheets(Original_Sheet)
    ActiveSheet.Name = Copy_Sheet

    Worksheets(Copy_Sheet).Outline.ShowLevels RowLevels:=1

End Function
Public Function Backup_Offset(Original_Sheet As String, Copy_Sheet As String)
    Sheets(Original_Sheet).Select
    Rows("4:1000").Select
    Selection.Copy
    Sheets(Copy_Sheet).Select
    ActiveSheet.Range("A4").Select
    ActiveSheet.Paste
    Sheets(Original_Sheet).Select
End Function
Public Function Output_Offset(Sheets As String, mode As Integer)
'>>Description (Mode)
'   Mode = 0 ... Offset Cell All Clear Mode
'   Mode = 1 ... Offset Write Mode

    '==================================
    
    Dim Start_cell_x As Long    'Read start Cell X adress
    Dim Start_cell_y As Long    'Read start Cell Y adress
    Dim Cell_count_x As Long    'Count X Cell
    Dim Cell_count_y As Long    'Count Y Cell
    Dim Counter_Cond    As Integer
    Dim Counter_Pin     As Integer
    Dim Counter_Site    As Integer
    Dim Counter_Node    As Integer
    Dim write_Count As Integer
    Dim Cell_Count As Integer
    '==================================
    
    Dim start_Site As Integer
    Dim start_Condition As Integer
    Dim start_PinName As Integer
    Dim end_Site As Integer
    Dim end_Condition As Integer
    Dim end_PinName As Integer

    'Defult Config ====================
    start_Site = 0
    start_Condition = 0
    start_PinName = 0
    end_Site = nSite
    end_Condition = Condition_Number - 1
    end_PinName = pin_number - 1
    '==================================

    'On Lock Flag =====================
    If Flg_SiteLock = -1 Then
        start_Site = frm_PALS_VoltAdj_Main.SitePrint.ListIndex
        end_Site = frm_PALS_VoltAdj_Main.SitePrint.ListIndex
    End If
    If Flg_ConditionLock = -1 Then
        start_Condition = frm_PALS_VoltAdj_Main.CONDITIONPrint.ListIndex
        end_Condition = frm_PALS_VoltAdj_Main.CONDITIONPrint.ListIndex
    End If
    If Flg_PinLock = -1 Then
        start_PinName = frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex
        end_PinName = frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex
    End If
    '==================================

    'Write Offset Data ==============================================
    Start_cell_x = Start_x + 4
    Start_cell_y = Start_y + 2
    Cell_count_x = 0
    Cell_count_y = 0
    Cell_Count = 0
    write_Count = 0
    
    For Counter_Node = 0 To Tester_exist_cnt - 1
        If Tester_Name(Counter_Node) = Sw_Node Then
            Cell_count_x = 0
'            ReDim Preserve Offset_Voltage_vc(Condition_Number - 1, Pin_Number - 1, Site_Number(Counter_Node)) 'Dynamic Array
            
            For Counter_Pin = start_PinName To end_PinName
                Cell_count_y = 0
                If Flg_PinLock = -1 Then Cell_count_x = Cell_count_x + Counter_Pin
                
                For Counter_Site = start_Site To end_Site
                    If Flg_SiteLock = -1 Then Cell_count_y = Cell_count_y + start_Site
                    Start_cell_y = Next_Tester_YAddress(Counter_Node)
                                        
                    For Counter_Cond = start_Condition To end_Condition
                    If Flg_ConditionLock = -1 Then Start_cell_y = Start_cell_y + Next_Condition_YAddress * Counter_Cond

'                        If Cell_Count >= Break_Counter Then Exit Function
                        If mode = 0 Then        'Cell Clear Mode ==============
                            If Flg_Break_voltcheck = False _
                            Or (Flg_Break_voltcheck = True And write_Count > Break_Counter) Then
                                Worksheets(Sheets).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x).Value = ""
'                                write_Count = write_Count + 1
                            End If
                        
                        ElseIf mode = 1 Then    'Offset Write Mode ============
                        
                            If (Flg_PinLock = 0 Or (Flg_PinLock = -1 And Counter_Pin = frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex)) _
                            And (Flg_SiteLock = 0 Or (Flg_SiteLock = -1 And Counter_Site = frm_PALS_VoltAdj_Main.SitePrint.ListIndex)) _
                            And (Flg_ConditionLock = 0 Or (Flg_ConditionLock = -1 And Counter_Cond = frm_PALS_VoltAdj_Main.CONDITIONPrint.ListIndex)) Then
                            If CheckRun(Counter_Pin) = "ON" Then
                                
                                If (Flg_Break_voltcheck = False And (Flg_Stop_voltcheck = False Or (Flg_Stop_voltcheck = True And Cell_Count < Break_Counter))) _
                                Or (Flg_Break_voltcheck = True And Cell_Count >= Break_Counter_prev And Cell_Count < Break_Counter) Then
                                    'Write Offset
                                    Worksheets(Sheets).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x).Value = Offset_Voltage_vc(Counter_Cond, Counter_Pin, Counter_Site)
                                    'Color=Cyan
                                    Worksheets(Sheets).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x).Interior.ColorIndex = 8
                                    'Color=Yellow(CheckPin error)
                                    If CheckPin_flag(Counter_Cond, Counter_Pin, Counter_Site) = 1 Then Worksheets(Sheets).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x).Interior.ColorIndex = 6
                                    'Color=Red(Limit Over)
                                    If Abs(val(Offset_Voltage_vc(Counter_Cond, Counter_Pin, Counter_Site))) >= LimitOffset Then Worksheets(Sheets).Cells(Start_cell_y + Cell_count_y, Start_cell_x + Cell_count_x).Interior.ColorIndex = 3

                                End If
                                
'                                write_Count = write_Count + 1
                            End If
                            Cell_Count = Cell_Count + 1
                            End If
                        End If
                        
                        Start_cell_y = Start_cell_y + Next_Condition_YAddress
                    Next Counter_Cond
                    
                    Cell_count_y = Cell_count_y + 1
                Next Counter_Site
                
                Cell_count_x = Cell_count_x + 1
            Next Counter_Pin
        
        End If
    Next Counter_Node
    '================================================================

End Function

Public Sub Clear_Navi(Optional Flg_voltcheck_end As Integer = 0)

    'For Clear Button
    If Flg_Clear_voltcheck = True Then
        frm_PALS_VoltAdj_Main.SiteLockCheck.Value = 0
        frm_PALS_VoltAdj_Main.ConditionLockCheck.Value = 0
        frm_PALS_VoltAdj_Main.PinLockCheck.Value = 0
        frm_PALS_VoltAdj_Main.SitePrint.ListIndex = -1
        frm_PALS_VoltAdj_Main.CONDITIONPrint.ListIndex = -1
        frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex = -1
        frm_PALS_VoltAdj_Main.PinNumPrint.Value = ""
    End If

    'Voltage Clear
    If Flg_SiteLock = -1 Or Flg_PinLock = -1 Or Flg_ConfirmMode = -1 Then
    Else
        frm_PALS_VoltAdj_Main.SetupVolPrint.Value = ""
        frm_PALS_VoltAdj_Main.OffsetVolPrint.Value = ""
        frm_PALS_VoltAdj_Main.InputVolPrint.Value = ""
        frm_PALS_VoltAdj_Main.MeasVolPrint.Value = ""
        frm_PALS_VoltAdj_Main.DiffVolPrint.Value = ""
    End If
    
    'LockCheck Clear On Not Select Status
    If Flg_SiteLock = -1 And frm_PALS_VoltAdj_Main.SitePrint.ListIndex = -1 Then
        frm_PALS_VoltAdj_Main.SiteLockCheck.Value = 0
    End If
    If Flg_PinLock = -1 And frm_PALS_VoltAdj_Main.PinNamePrint.ListIndex = -1 Then
        frm_PALS_VoltAdj_Main.PinLockCheck.Value = 0
    End If
    If Flg_ConfirmMode = -1 And frm_PALS_VoltAdj_Main.CONDITIONPrint.ListIndex = -1 Then
        frm_PALS_VoltAdj_Main.ConditionLockCheck.Value = 0
    End If

''    'Flow Clear
''    frm_PALS_VoltAdj_Main.ConnectCheck.Visible = False
''    frm_PALS_VoltAdj_Main.Applycheck.Visible = False
''    frm_PALS_VoltAdj_Main.OffsetCheck.Visible = False
    
    'Button Visiable
    frm_PALS_VoltAdj_Main.ClearButton.enabled = True
    frm_PALS_VoltAdj_Main.BreakButton.Visible = False
    frm_PALS_VoltAdj_Main.RunButton.Visible = True
    frm_PALS_VoltAdj_Main.ClearButton.Visible = True
    frm_PALS_VoltAdj_Main.ExitButton.Visible = True
    
    'ProgressBar Clear
    frm_PALS_VoltAdj_Main.progress_lbl.width = 1
    frm_PALS_VoltAdj_Main.ProgressPer = 0
    
    'Option Clear
    If Flg_voltcheck_end = 0 Then
        frm_PALS_VoltAdj_Main.ConfirmModeCheck.Value = 0
        frm_PALS_VoltAdj_Main.DataLogCheck.Value = 0
        frm_PALS_VoltAdj_Main.Onlyvoltagecheck.Value = 0
    End If
    
    'Flag Clear
    Flg_Stop_voltcheck = False
    Flg_Break_voltcheck = False
    
    'Return Message On Open Navigation
    frm_PALS_VoltAdj_Main.Message = "[Run]   -> Run VoltCheck" + vbCrLf + _
                            "if you want to Check opptional,Select[Pin]&Check[Lock]" + vbCrLf + vbCrLf + _
                            "[Clear] -> Clear All Check Conditon"
                            
    DoEvents

End Sub




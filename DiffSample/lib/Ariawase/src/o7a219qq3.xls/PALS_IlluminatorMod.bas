Attribute VB_Name = "PALS_IlluminatorMod"
'*********************************************************************************************************
' DESIGN FOR JOB STANDARD PROGRAM
'
' Revision History:
'   Release Ver1.00 2006/07/24 S.Fukumaru Production Release
'   Release Ver2.00 2010/10/21 M.Imamura  + csOptCond + csOptCondParams
'   Release Ver3.00 2011/06/13 M.Imamura  + Mcommand taiou
'                                           Opt_Judgement add.
'                                           sub_MakeLuxTable Set_LuxToMcommand add.
'                                           OptSet_NIKON mod.
'   Release Ver3.10 2011/06/13 M.Imamura  NsisIni add.
'                                         Opt_Judgment/ReadOptLux mod.
'   Release Ver4.00 2011/12/12 M.Imamura  OptCheck add.
'                                         Flg_TOPTWait Add.
'   Release Ver4.10 2012/04/09 M.Imamura  Flg_TOPTWait Mod.
'   Release Ver5.00 2012/10/18 M.Imamura  OptModZ ->Sheet , Down2-Down6 Add. ,Flg_DefaultLB Add.
'           Ver5.01 2013/12/02 H.Arikawa  Debug F.B.
'**********************************************************************************************************

Option Explicit

Public OptCond As csPALS_OptCond

'============================
'=========For DEBUG==========
'============================
Public Flg_Illum_TimeMeasure As Long            '==> Measure Illuminator Time
Public Flg_Illum_Disable As Long                '==> Illuminator Kill
Public Flg_Illum_ControlCommandDisply As Long   '==> Display ControlCommand Log
Public Flg_Illum_AxisDisable As Long            '==> Disable Axis Commnad For Debug(NIKON)

'============================
'========= COMMON ===========
'============================
Public Const NIKON = "NIKON"
Public Const INTERACTION = "InterAction"
Public Const KESILLUM = "KES"

Private hOptPort As Integer                      'For GPIB Direct Command Send Addr

'============================
'========== NIKON ===========
'============================
Public NSIS_II As Illuminator
Public Const NIKON_WRKSHT_NAME = "Opt(NSIS)"
Public Flg_Down As Boolean

'==================================
'========== InterAction ===========
'==================================
Public Flg_Illum_GPIB_StatusRead_IA As Boolean
Public Const IA_WRKSHT_NAME = "Opt(IA)"
Private Const STATUS_READ_COMMAND = "Q"
Private Const INITIAL_COMMAND = "DC"    'TakeCare! WAKASA Comannd "DCL"

'>>>2011/6/13 M.IMAMURA Add.
Public Const g_blnUseCSV = True
Public g_blnOptAdjusting As Boolean
Private dblLuxTable(20, 3) As Double
Private Const LUX_ADJ_LIM = 0.005       '0.5%
Private Const LUX_ADJ_LIM_CNT = 15      '15 times Trial
Private Const DEFAULT_ILLUM_NSIS2 = "1" 'NIKON NSIS2
Private Const DEFAULT_ILLUM_NSIS3 = "2" 'NIKON NSIS3
Private Const DEFAULT_ILLUM_NSIS5 = "2" 'NIKON NSIS5

Public Const NOTAVAILABLE_CMD = "NotAvailable"

Private Fail_Count As Integer
Public ArmOpt As String
Public ArmLux As Double
Public Flg_Opt_Judgment As Boolean

Public Const g_blnOptMcomMode = False

'<<<2011/6/13 M.IMAMURA Add.
'>>>2011/8/8 M.IMAMURA Add.
Public blnAutoAxisSet As Boolean

'<<<2011/8/8 M.IMAMURA Add.

'>>>2011/12/12 M.IMAMURA Add.For OptCheck
Public Opt_Lux As Double
Public OptResult As Double

Private Const OptLimit_NSIS3 As Integer = 7400
Private Const OptLimit_NSIS5 As Integer = 12800

Private Const OptLimit_NSIS3LB_LOW As Integer = 7400
Private Const OptLimit_NSIS3LB_HIGH As Integer = 7400

Private Const OptLimit_NSIS5LB_LOW As Integer = 12800
Private Const OptLimit_NSIS5LB_HIGH As Integer = 12800

Private Const OptLimit_ERROR As Integer = 30000
Private Const LUX_ERROR_VALUE As Integer = -9999    'Error code. Case Return Error Lux.
'<<<2011/12/12 M.IMAMURA Add.

'>>>2012/3/13 M.IMAMURA Add.For Opt_judgment
Private Const dblOptJudgeLimit As Double = 0.035          '0.035=+-3.5%
'<<<2012/3/13 M.IMAMURA Add.For Opt_judgment

'>>>2012/10/18 M.IMAMURA Add.
Public Const Flg_DefaultLB As Boolean = False

Public Flg_DownPosi As DownPosi
Public Enum DownPosi
    DownPosi_B1 = 1
    DownPosi_Up = 2
    DownPosi_Down = 3
    DownPosi_Down2 = 4
    DownPosi_Down3 = 5
    DownPosi_Down4 = 6
    DownPosi_Down5 = 7
    DownPosi_Down6 = 8
End Enum

Public Flg_FUnit As FUnit
Public Enum FUnit
    FUnit_THROUGH = 1
    FUnit_PIN = 2
    FUnit_F_UNIT = 3
End Enum
'<<<2012/10/18 M.IMAMURA Add.
'<<<2012/12/11 H.ARIKAWA Add.
'=================================================
Public G_NsisCnd(100) As NSIS_II_CONDITION '2012/11/15 175Debug Arikawa
Public G_NsisCndStr(100) As String         '2012/11/15 175Debug Arikawa
Public Type NSIS_II_CONDITION
     Axis As Long
     Level As Double
     Slit As Long
     NDFilter As Long
     WedgeFilter As Long
     color As Long
     Shutter As Long
     LCShutter As Long
     Diffusion As Long
     Pattern As Long
     FNumberIris As Long
     Mirror As Long
     LED As Long
     Pupil As Long
     Illuminant As Long
     ColorTemperature As Long
     FNumberTurret As Long
End Type
'<<<2012/12/11 H.ARIKAWA Add.

'=============================
'========== COMMON ===========
'=============================
Public Sub OptIni()

    '>>>2011/6/13 M.IMAMURA Add.
    If g_blnUseCSV = True And g_blnOptAdjusting = False Then Call ReadOptFile
    '<<<2011/6/13 M.IMAMURA Add.

    '========= Illumnator Disable =========
    If Flg_Illum_Disable = 1 Or Flg_Simulator = 1 Then Exit Sub
'    Hsy_optnum = 4.9
    '========= Get OptInformation =========
    Set OptCond = Nothing
    Set OptCond = New csPALS_OptCond
    
    '========= Select Illuminator Maker =========
    '******* NIKON *******
    If OptCond.IllumMaker = NIKON Then
        Call NsisIni
        
        Set NSIS_II = TheIllum.Illuminators("N-SIS II")
        If OptCond.IllumModel = "N-SIS3KAI" Then
            Call OptMod("PIN")
        End If
        NSIS_II.Initialize
        Flg_FUnit = FUnit_PIN
        Flg_DownPosi = DownPosi_Up

        Call OptCheck
        Call OptStatus
        TheExec.Datalog.WriteComment "Illuminator is Auto Mode!!"
        TheExec.Datalog.WriteComment "Illuminator initializing ..."
        TheHdw.WAIT 1
        Call GetOptLevelJudgment ' Flg_Illum_AxisDisable (0:"Axis.Level" 1:"Level")
        Call OptIni_NIKON
        TheExec.Datalog.WriteComment "Illuminator initialize Finished!!"
    '******* InterAction *******
    ElseIf OptCond.IllumMaker = INTERACTION Or OptCond.IllumMaker = KESILLUM Then
        Flg_Illum_GPIB_StatusRead_IA = CheckGpibStatusFlg
        
        Call NsisIni
        
        '===== Initalize Illuminator =====
        If Flg_Illum_GPIB_StatusRead_IA = True And OptCond.IllumMaker = INTERACTION Then
            Call OptSend_GPIBCommand(INITIAL_COMMAND)  'For Illum Initialize
        End If
        '=================================

    '******* ERROR *******
    Else
        MsgBox "You Select Illegal IllumMaker![" & OptCond.IllumMaker & "]" & vbCrLf & "ErrCode.4-5-01-5-36", vbExclamation, "<<Error>>@Optini"
        OptCond.IllumMaker = "You Select Wrong Illuminator!"
        Exit Sub
    End If
    '=============================================

End Sub
'=============================
'========== COMMON ===========
'=============================
Public Sub OptSet_Axis(strIllumMode As String, Optional Flg_TOPTWait As Boolean = False)

    Dim start_time As Double
    Dim end_time As Double

    '========= Illumnator Disable =========
    If Flg_Illum_ControlCommandDisply = 1 Then TheExec.Datalog.WriteComment "OptSet Status(" & strIllumMode & ")"
    If Flg_Illum_Disable = 1 Or Flg_Simulator = 1 Then Exit Sub
    
    '############## Debug Time Monitor ################
    If Flg_Illum_TimeMeasure = 1 Then start_time = TheExec.timer(0)
    '##################################################
    
    '>>>2011/06/13 M.IMAMURA OptReset Add.
    Call sub_CheckOptCond
    '<<<2011/06/13 M.IMAMURA OptReset Add.

    If OptCond.IllumMaker = NIKON Then
        '******* NIKON *******
        Call OptSet_NIKON_Axis(strIllumMode, Flg_TOPTWait)

    ElseIf OptCond.IllumMaker = INTERACTION Or OptCond.IllumMaker = KESILLUM Then
        '******* InterAction *******
        Call OptSet_IA(strIllumMode)
    Else
        '******* ERROR *******
        MsgBox "You Select Illegal IllumMaker![" & OptCond.IllumMaker & "]" & vbCrLf & "ErrCode.4-5-02-5-37", vbExclamation, "<<Error>>@OptSet"
    End If

    '############## Debug Time Monitor ################
    If Flg_Illum_TimeMeasure = 1 Then
        end_time = TheExec.timer(start_time)
        TheExec.Datalog.WriteComment "Optset(" & strIllumMode & ") Time = " & Format(1000 * end_time, "###.###") & "mSec"
        TheExec.Datalog.WriteComment " "
    End If
    '##################################################

End Sub
'=============================
'========== COMMON ===========
'=============================
Public Sub OptSet_Device(strIllumMode As String, Optional Flg_TOPTWait As Boolean = False)

    Dim start_time As Double
    Dim end_time As Double

    '========= Illumnator Disable =========
    If Flg_Illum_ControlCommandDisply = 1 Then TheExec.Datalog.WriteComment "OptSet Status(" & strIllumMode & ")"
    If Flg_Illum_Disable = 1 Or Flg_Simulator = 1 Then Exit Sub
    
    '############## Debug Time Monitor ################
    If Flg_Illum_TimeMeasure = 1 Then start_time = TheExec.timer(0)
    '##################################################
    
    '>>>2011/06/13 M.IMAMURA OptReset Add.
    Call sub_CheckOptCond
    '<<<2011/06/13 M.IMAMURA OptReset Add.

    If OptCond.IllumMaker = NIKON Then
        '******* NIKON *******
        Call OptSet_NIKON_Device(strIllumMode, Flg_TOPTWait)

    ElseIf OptCond.IllumMaker = INTERACTION Or OptCond.IllumMaker = KESILLUM Then
        '******* InterAction *******
        Call OptSet_IA(strIllumMode)
    Else
        '******* ERROR *******
        MsgBox "You Select Illegal IllumMaker![" & OptCond.IllumMaker & "]" & vbCrLf & "ErrCode.4-5-02-5-37", vbExclamation, "<<Error>>@OptSet"
    End If

    '############## Debug Time Monitor ################
    If Flg_Illum_TimeMeasure = 1 Then
        end_time = TheExec.timer(start_time)
        TheExec.Datalog.WriteComment "Optset(" & strIllumMode & ") Time = " & Format(1000 * end_time, "###.###") & "mSec"
        TheExec.Datalog.WriteComment " "
    End If
    '##################################################

End Sub
'=============================
'========== COMMON ===========
'=============================
Public Sub OptSet_Test(strIllumMode As String, Optional Flg_TOPTWait As Boolean = False)

    Dim start_time As Double
    Dim end_time As Double

    '========= Illumnator Disable =========
    If Flg_Illum_ControlCommandDisply = 1 Then TheExec.Datalog.WriteComment "OptSet Status(" & strIllumMode & ")"
    If Flg_Illum_Disable = 1 Or Flg_Simulator = 1 Then Exit Sub
    
    '############## Debug Time Monitor ################
    If Flg_Illum_TimeMeasure = 1 Then start_time = TheExec.timer(0)
    '##################################################
    
    '>>>2011/06/13 M.IMAMURA OptReset Add.
    Call sub_CheckOptCond
    '<<<2011/06/13 M.IMAMURA OptReset Add.

    If OptCond.IllumMaker = NIKON Then
        '******* NIKON *******
        Call OptSet_NIKON_Test(strIllumMode, Flg_TOPTWait)

    ElseIf OptCond.IllumMaker = INTERACTION Or OptCond.IllumMaker = KESILLUM Then
        '******* InterAction *******
        Call OptSet_IA(strIllumMode)
    Else
        '******* ERROR *******
        MsgBox "You Select Illegal IllumMaker![" & OptCond.IllumMaker & "]" & vbCrLf & "ErrCode.4-5-02-5-37", vbExclamation, "<<Error>>@OptSet"
    End If

    '############## Debug Time Monitor ################
    If Flg_Illum_TimeMeasure = 1 Then
        end_time = TheExec.timer(start_time)
        TheExec.Datalog.WriteComment "Optset(" & strIllumMode & ") Time = " & Format(1000 * end_time, "###.###") & "mSec"
        TheExec.Datalog.WriteComment " "
    End If
    '##################################################

End Sub
'=============================
'========== COMMON ===========
'=============================
Public Sub OptSet(strIllumMode As String, Optional Flg_TOPTWait As Boolean = False)

    Dim start_time As Double
    Dim end_time As Double

    '========= Illumnator Disable =========
    If Flg_Illum_ControlCommandDisply = 1 Then TheExec.Datalog.WriteComment "OptSet Status(" & strIllumMode & ")"
    If Flg_Illum_Disable = 1 Or Flg_Simulator = 1 Then Exit Sub
    
    '############## Debug Time Monitor ################
    If Flg_Illum_TimeMeasure = 1 Then start_time = TheExec.timer(0)
    '##################################################
    
    '>>>2011/06/13 M.IMAMURA OptReset Add.
    Call sub_CheckOptCond
    '<<<2011/06/13 M.IMAMURA OptReset Add.

    If OptCond.IllumMaker = NIKON Then
        '******* NIKON *******
        Call OptSet_NIKON(strIllumMode, Flg_TOPTWait)

    ElseIf OptCond.IllumMaker = INTERACTION Or OptCond.IllumMaker = KESILLUM Then
        '******* InterAction *******
        Call OptSet_IA(strIllumMode)
    Else
        '******* ERROR *******
        MsgBox "You Select Illegal IllumMaker![" & OptCond.IllumMaker & "]" & vbCrLf & "ErrCode.4-5-02-5-37", vbExclamation, "<<Error>>@OptSet"
    End If

    '############## Debug Time Monitor ################
    If Flg_Illum_TimeMeasure = 1 Then
        end_time = TheExec.timer(start_time)
        TheExec.Datalog.WriteComment "Optset(" & strIllumMode & ") Time = " & Format(1000 * end_time, "###.###") & "mSec"
        TheExec.Datalog.WriteComment " "
    End If
    '##################################################

End Sub
'=============================
'========== COMMON ===========
'=============================
Public Sub OptMod(strIllumMode As String, Optional Flg_TOPTWait As Boolean = False)

    Dim start_time As Double
    Dim end_time As Double

    '========= Illumnator Disable =========
    If Flg_Illum_Disable = 1 Or Flg_Simulator = 1 Then Exit Sub
    
    '############## Debug Time Monitor ################
    If Flg_Illum_TimeMeasure = 1 Then start_time = TheExec.timer(0)
    '##################################################

    '>>>2011/06/13 M.IMAMURA OptReset Add.
    Call sub_CheckOptCond
    '<<<2011/06/13 M.IMAMURA OptReset Add.
    
    If OptCond.IllumMaker = NIKON Then
        '******* NIKON *******
        Call OptMod_NIKON(strIllumMode, Flg_TOPTWait)

    ElseIf OptCond.IllumMaker = INTERACTION Or OptCond.IllumMaker = KESILLUM Then
        '******* InterAction *******
        Call OptSend_GPIBCommand(strIllumMode) ' Take Care! ## GPIB Command Direct Sending ##

    Else
        '******* ERROR *******
        MsgBox "You Select Illegal IllumMaker![" & OptCond.IllumMaker & "]" & vbCrLf & "ErrCode.4-5-03-5-38", vbExclamation, "<<Error>>@OptMod"
    End If

    '############## Debug Time Monitor ################
    If Flg_Illum_TimeMeasure = 1 Then
        end_time = TheExec.timer(start_time)
        TheExec.Datalog.WriteComment "Optmod(" & strIllumMode & ") Time = " & Format(1000 * end_time, "###.###") & "mSec"
        TheExec.Datalog.WriteComment " "
    End If
    '##################################################

End Sub
'============================
'========== NIKON ===========
'============================
Public Sub Opt_Judgment_Test(strIllumMode As String)

    Dim Low_Limit As Double
    Dim High_Limit As Double
    Dim Test_Mode As Long
    Dim OptLim As Double
    Dim OptRef As Double
    Dim OptTemp As Double
    Dim Fail_Count_Limit As Integer
'    Dim intCmdOptModFnumberTurret As Integer

    On Error GoTo EndOptJudgment
'        If OptCond.CondInfo(strIllumMode).OptModFnumberTurret = -1 Then
'            intCmdOptModFnumberTurret = 1
'        Else
'            intCmdOptModFnumberTurret = OptCond.CondInfo(strIllumMode).OptModFnumberTurret
'        End If

    '>>>2012/3/14 M.IMAMURA Change
    If OptCond.CondInfo(strIllumMode).OptJudge = "ON" Then
        OptLim = OptCond.CondInfo(strIllumMode).AxisLevel * dblOptJudgeLimit
    ElseIf OptCond.CondInfo(strIllumMode).OptJudge = "OFF" Then
        Exit Sub
    Else
        MsgBox "Opt_Judgment@Opt(NSIS) is Not Defined Opt_Judgment[" & strIllumMode & "]" & vbCrLf & "ErrCode.4-5-16-5-64", vbExclamation, "<<Error>>@Opt_Judgment"
    End If
    
'    Select Case strIllumMode
'        Case "HLX20"
'            OptLim = 500
'        Case "SHHL"
'            OptLim = 17.8
'        Case "LL"
'            OptLim = 8.5
'        Case Else
'            Exit Sub
'    End Select
    '<<<2012/3/14 M.IMAMURA Change
    
    '>>>2011/8/26 M.IMAMURA Del.
'    Call OptSend_GPIBCommand("M1,1," & CStr(OptCond.CondInfo(strIllumMode).WedgeFilter) & ",1,1," & CStr(OptCond.CondInfo(strIllumMode).NDFilter) & "," & intCmdOptModFnumberTurret & ",1," & Opt_LastCmd(strIllumMode))
'    Call OptStatus
    '<<<2011/8/26 M.IMAMURA Del.

    OptTemp = ReadOptLux
    OptRef = OptCond.CondInfo(strIllumMode).AxisLevel
    '---------OptLux Limit Set-------------------
    Low_Limit = OptRef - OptLim       'Low Limit Set
    High_Limit = OptRef + OptLim      'High Limit Set
    Fail_Count_Limit = 1          'Judgment Count Limit Set
    '-------------------------------------
    
    '//////////////// SENS Naibu SPEC Check /////////////////
    If Flg_Opt_Judgment = False Then
        If OptTemp < Low_Limit Or OptTemp > High_Limit Then         'Spec Over Check
            Fail_Count = Fail_Count + 1                             'Spec Over n-Count
        End If
        If Fail_Count >= Fail_Count_Limit Then
            Flg_Opt_Judgment = True
            ArmOpt = strIllumMode
            ArmLux = OptTemp
            Fail_Count = 0
            Call sub_errPALS("Illuminator Opt_Judgment[" & strIllumMode & "] Limit Error! Axis=" & CStr(OptRef) & ",Lux=" & CStr(OptTemp), "4-5-16-3-65", Enm_ErrFileBank_LOCAL)
        End If
    End If
    '////////////////////////////////////////////////////////
    Exit Sub
EndOptJudgment:
    Call sub_errPALS("Illuminator Command Error Opt_Judgment[" & strIllumMode & "]", "4-5-16-4-60", Enm_ErrFileBank_LOCAL)
'    MsgBox "Illuminator Command Error Opt_Judgment[" & strIllumMode & "]" & vbCrLf & "ErrCode.4-5-16-4-60", vbExclamation, "<<Error>>@Opt_Judgment"

End Sub

'============================
'========== NIKON ===========
'============================
Public Sub Opt_Judgment(strIllumMode As String)

    Dim Low_Limit As Double
    Dim High_Limit As Double
    Dim Test_Mode As Long
    Dim OptLim As Double
    Dim OptRef As Double
    Dim OptTemp As Double
    Dim Fail_Count_Limit As Integer
'    Dim intCmdOptModFnumberTurret As Integer

    On Error GoTo EndOptJudgment
'        If OptCond.CondInfo(strIllumMode).OptModFnumberTurret = -1 Then
'            intCmdOptModFnumberTurret = 1
'        Else
'            intCmdOptModFnumberTurret = OptCond.CondInfo(strIllumMode).OptModFnumberTurret
'        End If

    '>>>2012/3/14 M.IMAMURA Change
    If OptCond.CondInfo(strIllumMode).OptJudge = "ON" Then
        OptLim = OptCond.CondInfo(strIllumMode).AxisLevel * dblOptJudgeLimit
    ElseIf OptCond.CondInfo(strIllumMode).OptJudge = "OFF" Then
        Exit Sub
    Else
        MsgBox "Opt_Judgment@Opt(NSIS) is Not Defined Opt_Judgment[" & strIllumMode & "]" & vbCrLf & "ErrCode.4-5-16-5-64", vbExclamation, "<<Error>>@Opt_Judgment"
    End If
    
'    Select Case strIllumMode
'        Case "HLX20"
'            OptLim = 500
'        Case "SHHL"
'            OptLim = 17.8
'        Case "LL"
'            OptLim = 8.5
'        Case Else
'            Exit Sub
'    End Select
    '<<<2012/3/14 M.IMAMURA Change
    
    '>>>2011/8/26 M.IMAMURA Del.
'    Call OptSend_GPIBCommand("M1,1," & CStr(OptCond.CondInfo(strIllumMode).WedgeFilter) & ",1,1," & CStr(OptCond.CondInfo(strIllumMode).NDFilter) & "," & intCmdOptModFnumberTurret & ",1," & Opt_LastCmd(strIllumMode))
'    Call OptStatus
    '<<<2011/8/26 M.IMAMURA Del.

    OptTemp = ReadOptLux
    OptRef = OptCond.CondInfo(strIllumMode).AxisLevel
    '---------OptLux Limit Set-------------------
    Low_Limit = OptRef - OptLim       'Low Limit Set
    High_Limit = OptRef + OptLim      'High Limit Set
    Fail_Count_Limit = 1          'Judgment Count Limit Set
    '-------------------------------------
    
    '//////////////// SENS Naibu SPEC Check /////////////////
    If Flg_Opt_Judgment = False Then
        If OptTemp < Low_Limit Or OptTemp > High_Limit Then         'Spec Over Check
            Fail_Count = Fail_Count + 1                             'Spec Over n-Count
        End If
        If Fail_Count >= Fail_Count_Limit Then
            Flg_Opt_Judgment = True
            ArmOpt = strIllumMode
            ArmLux = OptTemp
            Fail_Count = 0
            Call sub_errPALS("Illuminator Opt_Judgment[" & strIllumMode & "] Limit Error! Axis=" & CStr(OptRef) & ",Lux=" & CStr(OptTemp), "4-5-16-3-65", Enm_ErrFileBank_LOCAL)
        End If
    End If
    '////////////////////////////////////////////////////////
    Exit Sub
EndOptJudgment:
    Call sub_errPALS("Illuminator Command Error Opt_Judgment[" & strIllumMode & "]", "4-5-16-4-60", Enm_ErrFileBank_LOCAL)
'    MsgBox "Illuminator Command Error Opt_Judgment[" & strIllumMode & "]" & vbCrLf & "ErrCode.4-5-16-4-60", vbExclamation, "<<Error>>@Opt_Judgment"

End Sub
'============================
'========== NIKON ===========
'============================
Private Function GetOptLevelJudgment() As Long

    Dim wkshtObj As Object
    Dim basePoint As Variant
    Dim nodeRow As Long
    Dim nodeColumn As Long
    Dim cmdstr As String
    
    On Error GoTo EndOpt

    '======= WorkSheet Select ========
    Set wkshtObj = ThisWorkbook.Sheets(NIKON_WRKSHT_NAME)
         
    '======= WorkSheet ErrorProcess ========
    If IsEmpty(wkshtObj) Then
        MsgBox "Non[" & NIKON_WRKSHT_NAME & "]WorkSheet For Optset" & vbCrLf & "ErrCode.4-5-04-8-39", vbExclamation, "<<Error>>@GetLevelJudgment"
        Exit Function
    End If
    
    '======= Base Point Find ========
    Set basePoint = Worksheets(NIKON_WRKSHT_NAME).Range("A1:K100").Find("Sw_node")
    If basePoint Is Nothing Then
        MsgBox "Non KeyWord[" & NIKON_WRKSHT_NAME & "]WorkSheet For Find Sw_node" & vbCrLf & "ErrCode.4-5-04-8-40", vbExclamation, "<<Error>>@GetLevelJudgment"
        Exit Function
    End If
    
    '======= Search Node Start Addres Define ========
    nodeRow = basePoint.Row + 1
    nodeColumn = basePoint.Column + 2

    '======= Search Node ========
    Do While Not wkshtObj.Cells(nodeRow, nodeColumn) = vbNullString
        If wkshtObj.Cells(nodeRow, nodeColumn) Like "*Level*" Then
            If Flg_Illum_AxisDisable = 1 Then
                cmdstr = "Level"
                If wkshtObj.Cells(nodeRow, nodeColumn) <> "Level" Then
                    wkshtObj.Cells(nodeRow, nodeColumn) = "Level"
                End If
                If Flg_Illum_ControlCommandDisply = 1 Then TheExec.Datalog.WriteComment "Optini Level Setting Direct Command"
            Else
                cmdstr = "Axis.Level"
                If wkshtObj.Cells(nodeRow, nodeColumn) <> "Axis.Level" Then
                     wkshtObj.Cells(nodeRow, nodeColumn) = "Axis.Level"
                End If
                If Flg_Illum_ControlCommandDisply = 1 Then TheExec.Datalog.WriteComment "Optini Level Setting Use Axis"
            End If
        End If
        nodeColumn = nodeColumn + 1
    Loop

    If cmdstr <> vbNullString Then Exit Function

EndOpt:
    MsgBox "Illuminator Command Error@Getting Level For Optini" & vbCrLf & "ErrCode.4-5-04-0-41", vbSystemModal, "<<Error>>@GetLevelJudgment"

End Function
'=============================
'========== COMMON ===========
'=============================
Public Sub OptSend_GPIBCommand(gpibCmd As String)

    Dim sndCmd As String
    On Error GoTo EndOpt

    '>>>2011/07/29 M.IMAMURA OptReset Add.
    If hOptPort = 0 Then
        Call NsisIni
    End If
    '<<<2011/07/29 M.IMAMURA OptReset Add.

    '========= Illumnator Disable =========
    If Flg_Illum_Disable = 1 Then Exit Sub
    
    '############ For Debug Print ##########
    If Flg_Illum_ControlCommandDisply = 1 Then TheExec.Datalog.WriteComment "SendGPIB Command (" & gpibCmd & ")"
    '#######################################

    If Trim(gpibCmd) = vbNullString Then
        MsgBox "gpibCmd is Empty" & vbCrLf & "ErrCode.4-5-05-4-42", vbSystemModal, "<<Error>>@OptSend_GPIBCommand"
        Exit Sub
    End If
    
    sndCmd = gpibCmd + Chr(13) + Chr(10)
    Call ibwrt(hOptPort, sndCmd)
    TheHdw.WAIT 0.02

    '=========== Status Read ============
    If OptCond.IllumMaker = NIKON Then
        Call ReadSRQ
        TheHdw.WAIT 0.02
    '>>> 2011/5/20 M.Imamura KESILLUM NotUse OptStatus
    ElseIf OptCond.IllumMaker = INTERACTION Then
        Call OptStatus
    End If
    '====================================
    
    Exit Sub
    
EndOpt:
    Call sub_errPALS("Illegal Parameter OptSend_GPIBCommand [" & gpibCmd & "]", "4-5-05-0-43", Enm_ErrFileBank_LOCAL)
'    MsgBox "Illegal Parameter OptSend_GPIBCommand [" & gpibCmd & "]" & vbCrLf & "ErrCode.4-5-05-0-43", vbExclamation, "<<Error>>@OptSend_GPIBCommand"

End Sub
'============================
'========== NIKON ===========
'============================
Public Sub OptModZ_NSIS5(strIllumMode As String, Optional Flg_TOPTWait As Boolean = False)

    '<< This Command Design For Only N-SIS5 PinHoleUnit Z Control>>

    Dim optcmd As String
    Dim strIllumModZ As String
    Dim OrgUnit As FUnit
    
    '========= Illumnator Disable =========
    If Flg_Illum_Disable = 1 Or Flg_Simulator = 1 Then Exit Sub
    
    Select Case strIllumMode
        Case "B1", "Up", "Down", "Down2", "Down3", "Down4", "Down5", "Down6":
            strIllumModZ = strIllumMode
        Case Else
            If OptCond.CondInfo(strIllumMode).OptModDownPosition = 1 Then
                strIllumModZ = "B1"
            ElseIf OptCond.CondInfo(strIllumMode).OptModDownPosition = 2 Then
                strIllumModZ = "Up"
            ElseIf OptCond.CondInfo(strIllumMode).OptModDownPosition = 3 Then
                strIllumModZ = "Down"
            ElseIf OptCond.CondInfo(strIllumMode).OptModDownPosition = 4 Then
                strIllumModZ = "Down2"
            ElseIf OptCond.CondInfo(strIllumMode).OptModDownPosition = 5 Then
                strIllumModZ = "Down3"
            ElseIf OptCond.CondInfo(strIllumMode).OptModDownPosition = 6 Then
                strIllumModZ = "Down4"
            ElseIf OptCond.CondInfo(strIllumMode).OptModDownPosition = 7 Then
                strIllumModZ = "Down5"
            ElseIf OptCond.CondInfo(strIllumMode).OptModDownPosition = 8 Then
                strIllumModZ = "Down6"
            End If
    End Select

    '>>>2011/06/13 M.IMAMURA OptReset Add.
    Call sub_CheckOptCond
    '<<<2011/06/13 M.IMAMURA OptReset Add.

    If OptCond.IllumMaker = NIKON Then
    
        '############ For Debug Print ##########
        If Flg_Illum_ControlCommandDisply = 1 Then TheExec.Datalog.WriteComment "OptMod Z Position(N-SIS5) Status(" & strIllumMode & ")"
        '#######################################
        
        Select Case strIllumModZ
            Case "B1":
                If Flg_DownPosi <> DownPosi_B1 Then
                    'Set to PIN
                    If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS3KAI" Then
                        If Flg_FUnit = FUnit_THROUGH Then
                            OrgUnit = FUnit_THROUGH
                            Call OptMod("PIN")
                        ElseIf Flg_FUnit = FUnit_F_UNIT Then
                            MsgBox "Illuminator Command Error DownPosition:B1-FnumberTurret:F_UNIT " & vbCrLf & "ErrCode.4-5-06-5-63", vbExclamation, "<<Error>>OptModZ_NSIS5"
                            Exit Sub
                        End If
                    ElseIf OptCond.IllumModel = "N-SIS5" Or OptCond.IllumModel = "N-SIS5KAI" Then
                        If Flg_FUnit = FUnit_PIN Then
                            MsgBox "Illuminator Command Error DownPosition:B1-FnumberTurret:PIN " & vbCrLf & "ErrCode.4-5-06-5-63", vbExclamation, "<<Error>>OptModZ_NSIS5"
                            Exit Sub
                        ElseIf Flg_FUnit = FUnit_F_UNIT Then
                            MsgBox "Illuminator Command Error DownPosition:B1-FnumberTurret:F_UNIT " & vbCrLf & "ErrCode.4-5-06-5-63", vbExclamation, "<<Error>>OptModZ_NSIS5"
                            Exit Sub
                        End If
                    End If
                    
                    optcmd = "B1"
                        Call OptSend_GPIBCommand(optcmd)
                    If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS3KAI" Then
                        TheHdw.WAIT 200 * mS
                    End If
                    Flg_DownPosi = DownPosi_B1

                    Call OptStatus(Flg_TOPTWait)

                    'Return to Original
                    If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS3KAI" Then
                        If OrgUnit = FUnit_THROUGH Then
                            Call OptMod("THROUGH", Flg_TOPTWait)
                        End If
                    End If
                    
                End If
            Case "Up": ' Pin Up-Side  (-32mm)
                If Flg_DownPosi <> DownPosi_Up Then
                    'Set to PIN
                    If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS3KAI" Then
                        If Flg_FUnit = FUnit_THROUGH Then
                            OrgUnit = FUnit_THROUGH
                            Call OptMod("PIN")
                        ElseIf Flg_FUnit = FUnit_F_UNIT Then
                            OrgUnit = FUnit_F_UNIT
                            Call OptMod("PIN")
                        End If
                    End If
                    
                    optcmd = "B2"
                    Call OptSend_GPIBCommand(optcmd)
                    If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS3KAI" Then
                        TheHdw.WAIT 200 * mS
                    End If
                    Flg_DownPosi = DownPosi_Up

                    Call OptStatus(Flg_TOPTWait)

                    'Return to Original
                    If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS3KAI" Then
                        If OrgUnit = FUnit_THROUGH Then
                            Call OptMod("THROUGH", Flg_TOPTWait)
                        ElseIf OrgUnit = FUnit_F_UNIT Then
                            Call OptMod("F_UNIT", Flg_TOPTWait)
                        End If
                    End If
                    
                End If
            Case "Down": ' Pin Down-Side ( EPD 54mm)
                If Flg_DownPosi <> DownPosi_Down Then
                    'Set to PIN
                    If OptCond.IllumModel = "N-SIS3" Then
                        If Flg_FUnit = FUnit_THROUGH Then
                            MsgBox "Illuminator Command Error DownPosition:Down-FnumberTurret:THROUGH " & vbCrLf & "ErrCode.4-5-06-5-63", vbExclamation, "<<Error>>OptModZ_NSIS5"
                            Exit Sub
                        ElseIf Flg_FUnit = FUnit_F_UNIT Then
                            MsgBox "Illuminator Command Error DownPosition:Down-FnumberTurret:F_UNIT " & vbCrLf & "ErrCode.4-5-06-5-63", vbExclamation, "<<Error>>OptModZ_NSIS5"
                            Exit Sub
                        End If
                    ElseIf OptCond.IllumModel = "N-SIS3KAI" Then
                        If Flg_FUnit = FUnit_THROUGH Then
                            OrgUnit = FUnit_THROUGH
                            Call OptMod("PIN")
                        ElseIf Flg_FUnit = FUnit_F_UNIT Then
                            OrgUnit = FUnit_F_UNIT
                            Call OptMod("PIN")
                        End If
                    ElseIf OptCond.IllumModel = "N-SIS5" Or OptCond.IllumModel = "N-SIS5KAI" Then
                        If Flg_FUnit = FUnit_THROUGH Then
                            MsgBox "Illuminator Command Error DownPosition:Down-FnumberTurret:THROUGH " & vbCrLf & "ErrCode.4-5-06-5-63", vbExclamation, "<<Error>>OptModZ_NSIS5"
                            Exit Sub
                        End If
                    End If
                    
                    optcmd = "B3"
                    Call OptSend_GPIBCommand(optcmd)
                    If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS3KAI" Then
                        TheHdw.WAIT 200 * mS
                    End If
                    Flg_DownPosi = DownPosi_Down

                    Call OptStatus(Flg_TOPTWait)

                    'Return to Original
                    If OptCond.IllumModel = "N-SIS3KAI" Then
                        If OrgUnit = FUnit_THROUGH Then
                            Call OptMod("THROUGH", Flg_TOPTWait)
                        ElseIf OrgUnit = FUnit_F_UNIT Then
                            Call OptMod("F_UNIT", Flg_TOPTWait)
                        End If
                    End If
                
                End If
            Case "Down2": ' Pin Down-Side ( EPD 51.5mm)
'                If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS3KAI" Then
                    MsgBox "Illuminator Command Error DownPosition:Down2" & vbCrLf & "ErrCode.4-5-06-5-63", vbExclamation, "<<Error>>OptModZ_NSIS5"
                    Exit Sub
'                End If
'                If Flg_DownPosi <> DownPosi_Down2 Then
'                    optcmd = "B4"
'                    Call OptSend_GPIBCommand(optcmd)
'                    Flg_DownPosi = DownPosi_Down2
'                End If
            
            Case "Down3": ' Pin Down-Side ( EPD 38.5mm)
'                If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS3KAI" Then
                    MsgBox "Illuminator Command Error DownPosition:Down3" & vbCrLf & "ErrCode.4-5-06-5-63", vbExclamation, "<<Error>>OptModZ_NSIS5"
                    Exit Sub
'                End If
'                If Flg_DownPosi <> DownPosi_Down3 Then
'                    optcmd = "B5"
'                    Call OptSend_GPIBCommand(optcmd)
'                    Flg_DownPosi = DownPosi_Down3
'                End If
            
            Case "Down4": ' Pin Down-Side ( EPD 30.5mm)
'                If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS3KAI" Then
                    MsgBox "Illuminator Command Error DownPosition:Down4" & vbCrLf & "ErrCode.4-5-06-5-63", vbExclamation, "<<Error>>OptModZ_NSIS5"
                    Exit Sub
'                End If
'                If Flg_DownPosi <> DownPosi_Down4 Then
'                    optcmd = "B6"
'                    Call OptSend_GPIBCommand(optcmd)
'                    Flg_DownPosi = DownPosi_Down4
'                End If
            
            Case "Down5": ' FHL
'                If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS3KAI" Then
                    MsgBox "Illuminator Command Error DownPosition:Down5" & vbCrLf & "ErrCode.4-5-06-5-63", vbExclamation, "<<Error>>OptModZ_NSIS5"
                    Exit Sub
'                End If
'                If Flg_DownPosi <> DownPosi_Down5 Then
'                    optcmd = "B7"
'                    Call OptSend_GPIBCommand(optcmd)
'                    Flg_DownPosi = DownPosi_Down5
'                End If
            
            Case "Down6": ' FHL
'                If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS3KAI" Then
                    MsgBox "Illuminator Command Error DownPosition:Down6" & vbCrLf & "ErrCode.4-5-06-5-63", vbExclamation, "<<Error>>OptModZ_NSIS5"
                    Exit Sub
'                End If
'                If Flg_DownPosi <> DownPosi_Down6 Then
'                    optcmd = "B8"
'                    Call OptSend_GPIBCommand(optcmd)
'                    Flg_DownPosi = DownPosi_Down6
'                End If
            Case Else
                MsgBox "OptModZ Command Error[" & strIllumMode & "]" & vbCrLf & "ErrCode.4-5-06-4-44", vbSystemModal, "<<Error>>@OptModZ_NSIS5"
        End Select
    End If
    
End Sub
'============================
'========== NIKON ===========
'============================
'Set Last of M Command
Private Function Opt_LastCmd(strIllumMode As String) As String
    
    On Error GoTo EndOptLastCmd
    
    'Set FL=1
    If OptCond.IllumModel = "N-SIS" Then
        Opt_LastCmd = DEFAULT_ILLUM_NSIS2
    'Set IrisPos=2
    ElseIf OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS3KAI" Then
        If strIllumMode = "INIT" Then
            Opt_LastCmd = DEFAULT_ILLUM_NSIS3
        Else
            Opt_LastCmd = OptCond.CondInfo(strIllumMode).FNumberTurret
        End If
        If Opt_LastCmd = -1 Then
            Opt_LastCmd = DEFAULT_ILLUM_NSIS3
        End If
    'Set IrisPos=2
    ElseIf OptCond.IllumModel = "N-SIS5" Or OptCond.IllumModel = "N-SIS5KAI" Then
        If strIllumMode = "INIT" Then
            Opt_LastCmd = DEFAULT_ILLUM_NSIS5
        Else
            Opt_LastCmd = OptCond.CondInfo(strIllumMode).FNumberTurret
        End If
        If Opt_LastCmd = -1 Then
            Opt_LastCmd = DEFAULT_ILLUM_NSIS5
        End If
    End If

    Exit Function

EndOptLastCmd:
    MsgBox "Illuminator Command Error Opt_LastCmd[" & strIllumMode & "]" & vbCrLf & "ErrCode.4-5-15-3-59", vbExclamation, "<<Error>>@Opt_LastCmd"

End Function
'============================
'========== NIKON ===========
'============================
Private Sub OptSet_NIKON_Test(strIllumMode As String, Optional Flg_TOPTWait As Boolean = False)

    Dim intCmdShutter As Integer
    Dim intCmdPattern As Integer
    Dim intCmdColor As Integer
    Dim intCmdOptModFnumberTurret As Integer
    Dim intCmdColorTemperature As Integer
   
    If Flg_Illum_Disable = 1 Then Exit Sub

    On Error GoTo EndOptSetNIKON

    With OptCond.CondInfo(strIllumMode)
    
''        '>>>2011/8/26 M.IMAMURA del.
''        If .AxisLevel > 0 Then
''            Call Opt_Judgment(strIllumMode)
''        End If
''        '<<<2011/8/26 M.IMAMURA del.
        If g_blnOptMcomMode = False Then
            If .AxisRegisterAxis > 0 Then
            '%%%%%%%%%%% Use Axis Set %%%%%%%%%
                Call NSIS_II.SetAxis(.AxisRegisterAxis)
'                Call OptStatus
                If Flg_Illum_AxisDisable = 1 Then
                    TheHdw.WAIT 0.5
                End If
                '>>>2011/8/26 M.IMAMURA Move.
'                Call Opt_Judgment(strIllumMode) 'For CIS
                '<<<2011/8/26 M.IMAMURA Move.
                If .Slit <> -1 Or .NDFilter <> -1 Or .WedgeFilter <> -1 Or .color <> -1 Or .Shutter <> -1 Or _
                   .LCShutter <> -1 Or .Diffusion <> -1 Or .Pattern <> -1 Or .FNumberIris <> -1 Or .Mirror <> -1 Or _
                   .LED <> -1 Or .Pupil <> -1 Or .Illuminant <> -1 Or .ColorTemperature <> -1 Or .FNumberTurret Then
                    Call NSIS_II.SetDevices(-1, .Slit, .NDFilter, .WedgeFilter, .color, .Shutter, .LCShutter _
                         , .Diffusion, .Pattern, .FNumberIris, .Mirror, .LED, .Pupil, .Illuminant, .ColorTemperature, .FNumberTurret)
                    Call OptStatus(Flg_TOPTWait)
                End If
            Else
            '%%%%%%%%%%% NonUse Axis Set %%%%%%%%%
                Call NSIS_II.SetDevices(.AxisLevel, .Slit, .NDFilter, .WedgeFilter, .color, .Shutter, .LCShutter _
                 , .Diffusion, .Pattern, .FNumberIris, .Mirror, .LED, .Pupil, .Illuminant, .ColorTemperature, .FNumberTurret)
                Call OptStatus(Flg_TOPTWait)
            End If
        Else
            If .Shutter = -1 Then
                intCmdShutter = 1
            Else
                intCmdShutter = .Shutter
            End If
        
            If .Pattern = -1 Then
                intCmdPattern = 1
            Else
                intCmdPattern = .Pattern
            End If

            If .color = -1 Then
                intCmdColor = 1
            Else
                intCmdColor = .color
            End If
        
            If .OptModFnumberTurret = -1 Then
                intCmdOptModFnumberTurret = 1
            Else
                intCmdOptModFnumberTurret = .OptModFnumberTurret
            End If
        
            If .ColorTemperature = -1 Then
                intCmdColorTemperature = 1
            Else
                intCmdColorTemperature = .ColorTemperature
            End If
        
            '%%%%%%%%%%% M Command Send %%%%%%%%%
            Call OptSend_GPIBCommand("M1," & intCmdShutter & "," & .WedgeFilter & "," & intCmdPattern & "," & intCmdColor & "," & .NDFilter & "," & intCmdOptModFnumberTurret & "," & intCmdColorTemperature & "," & Opt_LastCmd(strIllumMode))
            Call OptStatus(Flg_TOPTWait)

        End If
    End With

    Exit Sub

EndOptSetNIKON:
    Call sub_errPALS("Illuminator Command Error Optset[" & strIllumMode & "]", "4-5-14-4-58", Enm_ErrFileBank_LOCAL)
'    MsgBox "Illuminator Command Error Optset[" & strIllumMode & "]" & vbCrLf & "ErrCode.4-5-14-4-58", vbExclamation, "<<Error>>@OptSet_NIKON"

End Sub

'============================
'========== NIKON ===========
'============================
Private Sub OptSet_NIKON_Axis(strIllumMode As String, Optional Flg_TOPTWait As Boolean = False)

    Dim intCmdShutter As Integer
    Dim intCmdPattern As Integer
    Dim intCmdColor As Integer
    Dim intCmdOptModFnumberTurret As Integer
    Dim intCmdColorTemperature As Integer
   
    If Flg_Illum_Disable = 1 Then Exit Sub

    On Error GoTo EndOptSetNIKON

    With OptCond.CondInfo(strIllumMode)
    
''        '>>>2011/8/26 M.IMAMURA del.
''        If .AxisLevel > 0 Then
''            Call Opt_Judgment(strIllumMode)
''        End If
''        '<<<2011/8/26 M.IMAMURA del.
        If g_blnOptMcomMode = False Then
            If .AxisRegisterAxis > 0 Then
            '%%%%%%%%%%% Use Axis Set %%%%%%%%%
                Call NSIS_II.SetAxis(.AxisRegisterAxis)
'                Call OptStatus
                If Flg_Illum_AxisDisable = 1 Then
                    TheHdw.WAIT 0.5
                End If
                '>>>2011/8/26 M.IMAMURA Move.
'                Call Opt_Judgment(strIllumMode) 'For CIS
                '<<<2011/8/26 M.IMAMURA Move.
'                If .Slit <> -1 Or .NDFilter <> -1 Or .WedgeFilter <> -1 Or .color <> -1 Or .Shutter <> -1 Or _
'                   .LCShutter <> -1 Or .Diffusion <> -1 Or .Pattern <> -1 Or .FNumberIris <> -1 Or .Mirror <> -1 Or _
'                   .LED <> -1 Or .Pupil <> -1 Or .Illuminant <> -1 Or .ColorTemperature <> -1 Or .FNumberTurret Then
'                    Call NSIS_II.SetDevices(-1, .Slit, .NDFilter, .WedgeFilter, .color, .Shutter, .LCShutter _
'                         , .Diffusion, .Pattern, .FNumberIris, .Mirror, .LED, .Pupil, .Illuminant, .ColorTemperature, .FNumberTurret)
'                    Call OptStatus(Flg_TOPTWait)
'                End If
            Else
            '%%%%%%%%%%% NonUse Axis Set %%%%%%%%%
                Call NSIS_II.SetDevices(.AxisLevel, .Slit, .NDFilter, .WedgeFilter, .color, .Shutter, .LCShutter _
                 , .Diffusion, .Pattern, .FNumberIris, .Mirror, .LED, .Pupil, .Illuminant, .ColorTemperature, .FNumberTurret)
                Call OptStatus(Flg_TOPTWait)
            End If
        Else
            If .Shutter = -1 Then
                intCmdShutter = 1
            Else
                intCmdShutter = .Shutter
            End If
        
            If .Pattern = -1 Then
                intCmdPattern = 1
            Else
                intCmdPattern = .Pattern
            End If

            If .color = -1 Then
                intCmdColor = 1
            Else
                intCmdColor = .color
            End If
        
            If .OptModFnumberTurret = -1 Then
                intCmdOptModFnumberTurret = 1
            Else
                intCmdOptModFnumberTurret = .OptModFnumberTurret
            End If
        
            If .ColorTemperature = -1 Then
                intCmdColorTemperature = 1
            Else
                intCmdColorTemperature = .ColorTemperature
            End If
        
            '%%%%%%%%%%% M Command Send %%%%%%%%%
            Call OptSend_GPIBCommand("M1," & intCmdShutter & "," & .WedgeFilter & "," & intCmdPattern & "," & intCmdColor & "," & .NDFilter & "," & intCmdOptModFnumberTurret & "," & intCmdColorTemperature & "," & Opt_LastCmd(strIllumMode))
            Call OptStatus(Flg_TOPTWait)

        End If
    End With

    Exit Sub

EndOptSetNIKON:
    Call sub_errPALS("Illuminator Command Error Optset[" & strIllumMode & "]", "4-5-14-4-58", Enm_ErrFileBank_LOCAL)
'    MsgBox "Illuminator Command Error Optset[" & strIllumMode & "]" & vbCrLf & "ErrCode.4-5-14-4-58", vbExclamation, "<<Error>>@OptSet_NIKON"

End Sub

'============================
'========== NIKON ===========
'============================
Private Sub OptSet_NIKON_Device(strIllumMode As String, Optional Flg_TOPTWait As Boolean = False)

    Dim intCmdShutter As Integer
    Dim intCmdPattern As Integer
    Dim intCmdColor As Integer
    Dim intCmdOptModFnumberTurret As Integer
    Dim intCmdColorTemperature As Integer
   
    If Flg_Illum_Disable = 1 Then Exit Sub

    On Error GoTo EndOptSetNIKON

    With OptCond.CondInfo(strIllumMode)
    
''        '>>>2011/8/26 M.IMAMURA del.
''        If .AxisLevel > 0 Then
''            Call Opt_Judgment(strIllumMode)
''        End If
''        '<<<2011/8/26 M.IMAMURA del.
        If g_blnOptMcomMode = False Then
            If .AxisRegisterAxis > 0 Then
            '%%%%%%%%%%% Use Axis Set %%%%%%%%%
'                Call NSIS_II.SetAxis(.AxisRegisterAxis)
''                Call OptStatus
'                If Flg_Illum_AxisDisable = 1 Then
'                    TheHdw.WAIT 0.5
'                End If
                '>>>2011/8/26 M.IMAMURA Move.
'                Call Opt_Judgment(strIllumMode) 'For CIS
                '<<<2011/8/26 M.IMAMURA Move.
                If .Slit <> -1 Or .NDFilter <> -1 Or .WedgeFilter <> -1 Or .color <> -1 Or .Shutter <> -1 Or _
                   .LCShutter <> -1 Or .Diffusion <> -1 Or .Pattern <> -1 Or .FNumberIris <> -1 Or .Mirror <> -1 Or _
                   .LED <> -1 Or .Pupil <> -1 Or .Illuminant <> -1 Or .ColorTemperature <> -1 Or .FNumberTurret Then
                    Call NSIS_II.SetDevices(-1, .Slit, .NDFilter, .WedgeFilter, .color, .Shutter, .LCShutter _
                         , .Diffusion, .Pattern, .FNumberIris, .Mirror, .LED, .Pupil, .Illuminant, .ColorTemperature, .FNumberTurret)
                    Call OptStatus(Flg_TOPTWait)
                End If
            Else
            '%%%%%%%%%%% NonUse Axis Set %%%%%%%%%
                Call NSIS_II.SetDevices(.AxisLevel, .Slit, .NDFilter, .WedgeFilter, .color, .Shutter, .LCShutter _
                 , .Diffusion, .Pattern, .FNumberIris, .Mirror, .LED, .Pupil, .Illuminant, .ColorTemperature, .FNumberTurret)
                Call OptStatus(Flg_TOPTWait)
            End If
        Else
            If .Shutter = -1 Then
                intCmdShutter = 1
            Else
                intCmdShutter = .Shutter
            End If
        
            If .Pattern = -1 Then
                intCmdPattern = 1
            Else
                intCmdPattern = .Pattern
            End If

            If .color = -1 Then
                intCmdColor = 1
            Else
                intCmdColor = .color
            End If
        
            If .OptModFnumberTurret = -1 Then
                intCmdOptModFnumberTurret = 1
            Else
                intCmdOptModFnumberTurret = .OptModFnumberTurret
            End If
        
            If .ColorTemperature = -1 Then
                intCmdColorTemperature = 1
            Else
                intCmdColorTemperature = .ColorTemperature
            End If
        
            '%%%%%%%%%%% M Command Send %%%%%%%%%
            Call OptSend_GPIBCommand("M1," & intCmdShutter & "," & .WedgeFilter & "," & intCmdPattern & "," & intCmdColor & "," & .NDFilter & "," & intCmdOptModFnumberTurret & "," & intCmdColorTemperature & "," & Opt_LastCmd(strIllumMode))
            Call OptStatus(Flg_TOPTWait)

        End If
    End With

    Exit Sub

EndOptSetNIKON:
    Call sub_errPALS("Illuminator Command Error Optset[" & strIllumMode & "]", "4-5-14-4-58", Enm_ErrFileBank_LOCAL)
'    MsgBox "Illuminator Command Error Optset[" & strIllumMode & "]" & vbCrLf & "ErrCode.4-5-14-4-58", vbExclamation, "<<Error>>@OptSet_NIKON"

End Sub

'============================
'========== NIKON ===========
'============================
Private Sub OptSet_NIKON(strIllumMode As String, Optional Flg_TOPTWait As Boolean = False)

    Dim intCmdShutter As Integer
    Dim intCmdPattern As Integer
    Dim intCmdColor As Integer
    Dim intCmdOptModFnumberTurret As Integer
    Dim intCmdColorTemperature As Integer
   
    If Flg_Illum_Disable = 1 Then Exit Sub

    On Error GoTo EndOptSetNIKON

    With OptCond.CondInfo(strIllumMode)
    
''        '>>>2011/8/26 M.IMAMURA del.
''        If .AxisLevel > 0 Then
''            Call Opt_Judgment(strIllumMode)
''        End If
''        '<<<2011/8/26 M.IMAMURA del.
        If g_blnOptMcomMode = False Then
            If .AxisRegisterAxis > 0 Then
            '%%%%%%%%%%% Use Axis Set %%%%%%%%%
                Call NSIS_II.SetAxis(.AxisRegisterAxis)
                Call OptStatus
                If Flg_Illum_AxisDisable = 1 Then
                    TheHdw.WAIT 0.5
                End If
                '>>>2011/8/26 M.IMAMURA Move.
                Call Opt_Judgment(strIllumMode) 'For CIS
                '<<<2011/8/26 M.IMAMURA Move.
                If .Slit <> -1 Or .NDFilter <> -1 Or .WedgeFilter <> -1 Or .color <> -1 Or .Shutter <> -1 Or _
                   .LCShutter <> -1 Or .Diffusion <> -1 Or .Pattern <> -1 Or .FNumberIris <> -1 Or .Mirror <> -1 Or _
                   .LED <> -1 Or .Pupil <> -1 Or .Illuminant <> -1 Or .ColorTemperature <> -1 Or .FNumberTurret Then
                    Call NSIS_II.SetDevices(-1, .Slit, .NDFilter, .WedgeFilter, .color, .Shutter, .LCShutter _
                         , .Diffusion, .Pattern, .FNumberIris, .Mirror, .LED, .Pupil, .Illuminant, .ColorTemperature, .FNumberTurret)
                    Call OptStatus(Flg_TOPTWait)
                End If
            Else
            '%%%%%%%%%%% NonUse Axis Set %%%%%%%%%
                Call NSIS_II.SetDevices(.AxisLevel, .Slit, .NDFilter, .WedgeFilter, .color, .Shutter, .LCShutter _
                 , .Diffusion, .Pattern, .FNumberIris, .Mirror, .LED, .Pupil, .Illuminant, .ColorTemperature, .FNumberTurret)
                Call OptStatus(Flg_TOPTWait)
            End If
        Else
            If .Shutter = -1 Then
                intCmdShutter = 1
            Else
                intCmdShutter = .Shutter
            End If
        
            If .Pattern = -1 Then
                intCmdPattern = 1
            Else
                intCmdPattern = .Pattern
            End If

            If .color = -1 Then
                intCmdColor = 1
            Else
                intCmdColor = .color
            End If
        
            If .OptModFnumberTurret = -1 Then
                intCmdOptModFnumberTurret = 1
            Else
                intCmdOptModFnumberTurret = .OptModFnumberTurret
            End If
        
            If .ColorTemperature = -1 Then
                intCmdColorTemperature = 1
            Else
                intCmdColorTemperature = .ColorTemperature
            End If
        
            '%%%%%%%%%%% M Command Send %%%%%%%%%
            Call OptSend_GPIBCommand("M1," & intCmdShutter & "," & .WedgeFilter & "," & intCmdPattern & "," & intCmdColor & "," & .NDFilter & "," & intCmdOptModFnumberTurret & "," & intCmdColorTemperature & "," & Opt_LastCmd(strIllumMode))
            Call OptStatus(Flg_TOPTWait)

        End If
    End With

    Exit Sub

EndOptSetNIKON:
    Call sub_errPALS("Illuminator Command Error Optset[" & strIllumMode & "]", "4-5-14-4-58", Enm_ErrFileBank_LOCAL)
'    MsgBox "Illuminator Command Error Optset[" & strIllumMode & "]" & vbCrLf & "ErrCode.4-5-14-4-58", vbExclamation, "<<Error>>@OptSet_NIKON"

End Sub

'==================================
'========== InterAction ===========
'==================================
Private Sub OptSet_IA(strIllumMode As String)

    Dim optcmd As String

    optcmd = Get_IlluminatorParameter_IA(strIllumMode)

    If optcmd = "-1" Then Exit Sub
    
    Call OptSend_GPIBCommand(optcmd)
    
End Sub
'============================
'========== NIKON ===========
'============================
Private Sub OptMod_NIKON(strIllumMode As String, Optional Flg_TOPTWait As Boolean = False)

    On Error GoTo EndOptMod
   
    If Flg_Illum_ControlCommandDisply = 1 Then TheExec.Datalog.WriteComment "OptMod Status(" & strIllumMode & ")"

    Select Case strIllumMode
        Case "INIT"
            NSIS_II.SetDevices , , , , , , , , , , , , , , , 0  ' INITIAL
            Flg_FUnit = FUnit_THROUGH
        Case "THROUGH"
            NSIS_II.SetDevices , , , , , , , , , , , , , , , 1  ' THROUGH
            Flg_FUnit = FUnit_THROUGH
        Case "F_UNIT"
            NSIS_II.SetDevices , , , , , , , , , , , , , , , 3  ' F1.2
            Flg_FUnit = FUnit_F_UNIT
        Case "PIN"
            NSIS_II.SetDevices , , , , , , , , , , , , , , , 2  ' PIN-HOLE
            Flg_FUnit = FUnit_PIN
        Case Else
            NSIS_II.SetDevices , , , , , , , , , , , , , , , OptCond.CondInfo(strIllumMode).OptModFnumberTurret  ' WorkSheet Parameter
            Flg_FUnit = OptCond.CondInfo(strIllumMode).OptModFnumberTurret
    End Select

    Call OptStatus(Flg_TOPTWait)

    Exit Sub
EndOptMod:
    Call sub_errPALS("Illegal Parameter OptMod_NIKON[" & strIllumMode & "]", "4-5-07-0-45", Enm_ErrFileBank_LOCAL)
'    MsgBox "Illegal Parameter OptMod_NIKON[" & strIllumMode & "]" & vbCrLf & "ErrCode.4-5-07-0-45", vbExclamation, "<<Error>>@OptMod_NIKON"

End Sub
'============================
'========== NIKON ===========
'============================
Private Function OptIni_NIKON()

    Dim cmdStatus As String
    Dim Optcnt As Integer

    If g_blnOptMcomMode = True Then
        Call sub_MakeLuxTable
    End If

    For Optcnt = 0 To OptCond.OptCondNum

        If g_blnOptMcomMode = False Then
            '======== Axis Command Send ========
            cmdStatus = Send_Axis_Cmd_NIKON(Optcnt)
        Else
            If OptCond.CondInfoI(Optcnt).AxisRegisterAxis <> -1 Then
                cmdStatus = Set_LuxToMcommand(Optcnt)
            End If
        End If
    Next

End Function
'============================
'========== NIKON ===========
'============================
Private Sub sub_MakeLuxTable()
    Dim intNowND As Integer

    TheExec.Datalog.WriteComment "Illuminator Making Lux-ND-Wegde Table ..."
    
    For intNowND = OptCond.IllumNdMax To OptCond.IllumNdMin Step -1
'            Call OptSend_GPIBCommand("M1,Shutter,WedgeFilter,Pattern,Color,NDFilter,Fnumber,ColorTemperature,Iris")
            'SetWedge Min & ReadLux
            Call OptSend_GPIBCommand("M1,1," & CStr(OptCond.IllumWedgeMin) & ",1,1," & CStr(intNowND) & ",1,1," & Opt_LastCmd("INIT"))
            Call OptStatus
            dblLuxTable(intNowND, 0) = ReadOptLux
            
            'SetWedge Max & ReadLux
            Call OptSend_GPIBCommand("M1,1," & CStr(OptCond.IllumWedgeMax) & ",1,1," & CStr(intNowND) & ",1,1," & Opt_LastCmd("INIT"))
            Call OptStatus
            dblLuxTable(intNowND, 1) = ReadOptLux
    Next

End Sub
'============================
'========== NIKON ===========
'============================
Private Function Set_LuxToMcommand(Optcnt As Integer) As Boolean

    Dim status As String
    Dim intNextWedge As Integer
    Dim intNowWedge As Integer
    Dim intPrevWedge As Integer
    Dim dblLuxNow As Double
    Dim dblLuxPrev As Double
    Dim intNowND As Integer
    Dim intOptRetryCnt As Integer

    Set_LuxToMcommand = False
    intOptRetryCnt = 1
 
    With OptCond.CondInfoI(Optcnt)
 
        '########## LuxTableから使用NDを探す NdMax光量最小->NdMin光量最大
        TheExec.Datalog.WriteComment "Illuminator Searching Lux-ND-Wegde[" & .OptIdentifier & "] ... LuxTarget : " & .AxisLevel
        For intNowND = OptCond.IllumNdMax To OptCond.IllumNdMin Step -1
            'if Wedgemax<Lux<WedgeMin then exit for
            If dblLuxTable(intNowND, 1) < .AxisLevel And dblLuxTable(intNowND, 0) > .AxisLevel Then
                '使用NDをセットして抜ける
                .NDFilter = intNowND
                Exit For
            End If
            '最大光量で再現が不可能だったらエラー
            If intNowND = OptCond.IllumNdMin Then
                TheExec.Datalog.WriteComment "            YourLux[" & .OptIdentifier & "," & .AxisLevel & "Lux] is Bigger than NSIS MaxLux:" & CStr(dblLuxTable(OptCond.IllumNdMin, 0)) & "!!!!!!!!!!(T-T)"
                Call sub_errPALS("YourLux[" & .OptIdentifier & "," & .AxisLevel & "Lux] is Bigger than NSIS MaxLux:" & CStr(dblLuxTable(OptCond.IllumNdMin, 0)), "4-5-13-5-54", Enm_ErrFileBank_LOCAL)
'                MsgBox "YourLux[" & .OptIdentifier & "," & .AxisLevel & "Lux] is Bigger than NSIS MaxLux:" & CStr(dblLuxTable(OptCond.IllumNdMin, 0)) & vbCrLf & "ErrCode.4-5-13-5-54", vbExclamation, "<<Error>>Set_LuxToMcommand"
                Exit Function
            End If
        Next
        
        '########## Wedgeを変化させて、ターゲットに近づける
        '初期設定
        intNowWedge = OptCond.IllumWedgeMax
        intPrevWedge = OptCond.IllumWedgeMin
        dblLuxNow = dblLuxTable(intNowND, 1)
        dblLuxPrev = dblLuxTable(intNowND, 0)
        TheExec.Datalog.WriteComment "Illuminator Searching Lux-ND-Wegde[" & .OptIdentifier & "] ... ND : " & CStr(intNowND) & ",Wegde :" & CStr(intNextWedge) & ",Lux :" & CStr(dblLuxNow)
    
OptRetry:
        '次のウェッジを算出
        intNextWedge = Int(0.5 + intNowWedge _
                    + (intNowWedge - intPrevWedge) _
                    / (dblLuxNow - dblLuxPrev) _
                    * (.AxisLevel - dblLuxNow))
        'Mコマンドで制御
        Call OptSend_GPIBCommand("M1,1," & CStr(intNextWedge) & ",1,1," & CStr(intNowND) & ",1,1," & Opt_LastCmd("INIT"))
        Call OptStatus
        
        '光量取得
        dblLuxTable(intNowND, 2) = ReadOptLux
        
        intPrevWedge = intNowWedge
        intNowWedge = intNextWedge
        dblLuxPrev = dblLuxNow
        dblLuxNow = dblLuxTable(intNowND, 2)
        TheExec.Datalog.WriteComment "Illuminator Searching Lux-ND-Wegde[" & .OptIdentifier & "] ... ND : " & CStr(intNowND) & ",Wegde :" & CStr(intNextWedge) & ",Lux :" & CStr(dblLuxNow)
    
        'ターゲットと同等だったら終了
        If dblLuxNow = .AxisLevel Then
            Set_LuxToMcommand = True
        'ウェッジの動作ステップが1になったら
        ElseIf Abs(intNowWedge - intPrevWedge) = 1 Then
            'if Lux<Target*LUX_ADJ_LIM then exit
            If Abs(.AxisLevel - dblLuxNow) < .AxisLevel * LUX_ADJ_LIM Or Abs(.AxisLevel - dblLuxPrev) < .AxisLevel * LUX_ADJ_LIM Then
                Set_LuxToMcommand = True
                '前のウェッジのほうが差分が小さかったら戻す
                If Abs(.AxisLevel - dblLuxPrev) < Abs(.AxisLevel - dblLuxNow) Then
                    intNowWedge = intPrevWedge
                    dblLuxNow = dblLuxPrev
                End If
            Else
                TheExec.Datalog.WriteComment "            Searching Failed(T-T)!!!!!!!!!!"
                Call sub_errPALS("Illuminator Search Failed[" & .OptIdentifier & "] SearchStep=1... ", "4-5-13-5-55", Enm_ErrFileBank_LOCAL)
'                MsgBox "Illuminator Search Failed[" & .OptIdentifier & "] SearchStep=1... " & vbCrLf & "ErrCode.4-5-13-5-55", vbExclamation, "<<Error>>Set_LuxToMcommand"
            End If
        'Luxが変化しなかったら
        ElseIf dblLuxNow = dblLuxPrev Then
            If Abs(.AxisLevel - dblLuxNow) < .AxisLevel * LUX_ADJ_LIM Then
                Set_LuxToMcommand = True
            Else
                TheExec.Datalog.WriteComment "            Searching Failed(T-T)!!!!!!!!!!"
                Call sub_errPALS("Illuminator Search Failed[" & .OptIdentifier & "] Lux don't swing... ", "4-5-13-5-56", Enm_ErrFileBank_LOCAL)
'                MsgBox "Illuminator Search Failed[" & .OptIdentifier & "] Lux don't swing... " & vbCrLf & "ErrCode.4-5-13-5-56", vbExclamation, "<<Error>>Set_LuxToMcommand"
            End If
        '試行回数がLUX_ADJ_LIM_CNTに達したら
        ElseIf intOptRetryCnt = LUX_ADJ_LIM_CNT Then
            TheExec.Datalog.WriteComment "            Searching Failed(T-T)!!!!!!!!!!"
            Call sub_errPALS("Illuminator Search Failed[" & .OptIdentifier & "] AdjustCountOver[" & LUX_ADJ_LIM_CNT & "]... ", "4-5-13-3-57", Enm_ErrFileBank_LOCAL)
'            MsgBox "Illuminator Search Failed[" & .OptIdentifier & "] AdjustCountOver[" & LUX_ADJ_LIM_CNT & "]... " & vbCrLf & "ErrCode.4-5-13-3-57", vbExclamation, "<<Error>>Set_LuxToMcommand"
        Else
        '再度調整
            intOptRetryCnt = intOptRetryCnt + 1
            GoTo OptRetry
        End If
        
        If Set_LuxToMcommand = True Then
            TheExec.Datalog.WriteComment "            Searching Succeeded(^-^) [" & .OptIdentifier & "] ND : " & CStr(intNowND) & ",Wegde :" & CStr(intNowWedge) & ",Lux :" & CStr(dblLuxNow)
            'Wedge Set To Class
            .WedgeFilter = intNowWedge
            'AxisRegisterAxis Reset to -1
            .AxisRegisterAxis = -1
        Else
            .OptIdentifier = .OptIdentifier & NOTAVAILABLE_CMD
        End If

    End With

End Function
'=============================
'========== COMMON ===========
'=============================
Public Function OptStatus(Optional Flg_TOPTWait As Boolean = False) As Long

    Dim buf As String * 100
    Dim sendcom As String
    Dim status As String
    Dim Count As Integer
    
    Const STATUS_READY As String = "0"
    Const STATUS_BUSY As String = "1"
    Const STATUS_LOCAL As String = "2"
    Const STATUS_ERROR As String = "E"
    
    Dim iStatus As Long

    '========= Illumnator Disable =========
    If Flg_Illum_Disable = 1 Or Flg_Simulator = 1 Then Exit Function '2012/12/10 H.Arikawa
    
    '>>>2011/06/13 M.IMAMURA OptReset Add.
    Call sub_CheckOptCond
    '<<<2011/06/13 M.IMAMURA OptReset Add.

    '========= NIKON =========
    If OptCond.IllumMaker = NIKON And Flg_TOPTWait = False Then
        iStatus = NSIS_II.status
    
        While (iStatus <> 0)
            iStatus = NSIS_II.status
        Wend
    '=========== InterAction ===========
    ElseIf OptCond.IllumMaker = INTERACTION And Flg_Illum_GPIB_StatusRead_IA = True Then
    
        buf = Space$(100)
        sendcom = STATUS_READ_COMMAND + Chr(13) + Chr(10)
        status = STATUS_BUSY
        
        While (status <> STATUS_READY)
            '====== Status Read =====
            Call ibwrt(hOptPort, sendcom)
            Call ibrd(hOptPort, buf)
            status = Left(buf, 1)
            '========================
        
            If status = STATUS_BUSY Then
                Count = Count + 1
                If Count = 10000 Then
                    Call sub_errPALS("Illuminator is Busy Status.Timeout Error?", "4-5-08-9-46", Enm_ErrFileBank_LOCAL)
'                    MsgBox "Illuminator is Busy Status.Timeout Error?" & vbCrLf & "ErrCode.4-5-08-9-46", vbQuestion, "@OptStatus"
                    Stop
                    Exit Function
                End If
            End If
    
            If status = STATUS_LOCAL Then
                Call sub_errPALS("Illuminator is LOCAL Status", "4-5-08-9-47", Enm_ErrFileBank_LOCAL)
'                MsgBox "Illuminator is LOCAL Status" & vbCrLf & "ErrCode.4-5-08-9-47", vbExclamation, "<<Error>>@OptStatus"
                Stop
            End If
        
            If status = STATUS_ERROR Then
                Call sub_errPALS("Illuminator is Status Error", "4-5-08-9-48", Enm_ErrFileBank_LOCAL)
'                MsgBox "Illuminator is Status Error" & vbCrLf & "ErrCode.4-5-08-9-48", vbExclamation, "<<Error>>@OptStatus"
                Stop
            End If
        Wend
    End If
    
End Function
'============================
'========== NIKON ===========
'============================
Public Function ReadSRQ() As Integer

    Call ibrsp(hOptPort, ReadSRQ)

End Function


'============================
'========== NIKON ===========
'============================
Private Function Send_Axis_Cmd_NIKON(Optcnt As Integer) As String

    Dim flg_axisCmd As Boolean
    Dim status As String
    flg_axisCmd = False

    With NSIS_II
        '========== Axis Command Send ========
        '<<Pattern>>
        If OptCond.CondInfoI(Optcnt).Pattern <> -1 Then
            .Axis.Pattern = OptCond.CondInfoI(Optcnt).Pattern: flg_axisCmd = True
            status = status + "Pattern:" & OptCond.CondInfoI(Optcnt).Pattern & " "
        End If
        '<<Level>>
        If OptCond.CondInfoI(Optcnt).AxisLevel <> -1 Then
            .Axis.Level = OptCond.CondInfoI(Optcnt).AxisLevel: flg_axisCmd = True
            status = status + "Level:" & OptCond.CondInfoI(Optcnt).AxisLevel & " "
        End If
'        '<<Color>>
'        If OptCond.CondInfoI(Optcnt).Color <> -1 Then
'            .Axis.Color = OptCond.CondInfoI(Optcnt).Color: flg_axisCmd = True
'            status = status + "Color:" & OptCond.CondInfoI(Optcnt).Color & " "
'        End If
'        '<<NDFilter>>
'        If OptCond.CondInfoI(Optcnt).NDFilter <> -1 Then
'            .Axis.NDFilter = OptCond.CondInfoI(Optcnt).NDFilter: flg_axisCmd = True
'            status = status + "NDFilter:" & OptCond.CondInfoI(Optcnt).NDFilter & " "
'        End If
'        '<<WedgeFilter>>
'        If OptCond.CondInfoI(Optcnt).WedgeFilter <> -1 Then
'            .Axis.WedgeFilter = OptCond.CondInfoI(Optcnt).WedgeFilter: flg_axisCmd = True
'            status = status + "WedgeFilter:" & OptCond.CondInfoI(Optcnt).WedgeFilter & " "
'        End If
'        '<<Shutter>>
'        If OptCond.CondInfoI(Optcnt).Shutter <> -1 Then
'            .Axis.Shutter = OptCond.CondInfoI(Optcnt).Shutter: flg_axisCmd = True
'            status = status + "Shutter:" & OptCond.CondInfoI(Optcnt).Shutter & " "
'        End If
'        '<<ColorTemperature>>
'        If OptCond.CondInfoI(Optcnt).ColorTemperature <> -1 Then
'            .Axis.ColorTemp = OptCond.CondInfoI(Optcnt).ColorTemperature: flg_axisCmd = True
'            status = status + "ColorTemperature:" & OptCond.CondInfoI(Optcnt).ColorTemperature & " "
'        End If
'         '<<FNumberTurret>>
'        If OptCond.CondInfoI(Optcnt).FNumberTurret <> -1 Then
'            .Axis.F_Turret = OptCond.CondInfoI(Optcnt).FNumberTurret: flg_axisCmd = True
'            status = status + "FNumberTurret:" & OptCond.CondInfoI(Optcnt).FNumberTurret & " "
'        End If

        '========== AxisRegister Error Check ========
        '<<RegisterAxis>>
        If OptCond.CondInfoI(Optcnt).AxisRegisterAxis <> -1 Then
            If flg_axisCmd = True Then
                .RegisterAxis OptCond.CondInfoI(Optcnt).AxisRegisterAxis
                status = "RegisterAxis:" & Format(OptCond.CondInfoI(Optcnt).AxisRegisterAxis, "@@") & "  " & status
                '############ For Debug Print ##########
                If Flg_Illum_ControlCommandDisply = 1 And status <> vbNullString Then TheExec.Datalog.WriteComment "OptIni:" & status
                '#######################################
            Else
                MsgBox "Illegal Command@Optini" & vbCrLf & "ErrCode.4-5-09-5-49"
            End If
        Else
            flg_axisCmd = False
        End If
    End With

    '========== Status Check ========
    If flg_axisCmd = True Then
        Call OptStatus
        TheHdw.WAIT 1
    End If
    '================================

    Send_Axis_Cmd_NIKON = status

End Function

'==================================
'========== InterAction ===========
'==================================
Public Function Get_IlluminatorParameter_IA(strIllumMode As String) As String

    Dim cmdArgstr As String
    cmdArgstr = ""

    With OptCond.CondInfo(strIllumMode)
    
        'Command  S
        If .Shutter <> -1 Then cmdArgstr = cmdArgstr & "S" & .Shutter
        
        'Command  H
        If .Frosted <> -1 Then cmdArgstr = cmdArgstr & "H" & .Frosted
        
        'Command  L
        If .AxisLevel <> -1 Then cmdArgstr = cmdArgstr & "L" & .AxisLevel
        
        'Command  C
        If .color <> -1 Then cmdArgstr = cmdArgstr & "C" & .color
        
        'Command  N
        If .NDFilter <> -1 Then cmdArgstr = cmdArgstr & "N" & .NDFilter
        
        'Command  A
        If .WedgeFilter <> -1 Then cmdArgstr = cmdArgstr & "A" & .WedgeFilter
        
        'Command  P
        If .Pattern <> -1 Then cmdArgstr = cmdArgstr & "P" & .Pattern
        
        'Command  I
        If .EPD1 <> -1 Then cmdArgstr = cmdArgstr & "I" & .EPD1
        
        'Command  J
        If .EPD2 <> -1 Then cmdArgstr = cmdArgstr & "J" & .EPD2
    
        'Command  Z
        If .IrisPos <> -1 Then cmdArgstr = cmdArgstr & "Z" & .IrisPos
    
        'Command  X
        If .DeviceX <> -9999 Then cmdArgstr = cmdArgstr & "X" & .DeviceX
    
        'Command  Y
        If .DeviceY <> -9999 Then cmdArgstr = cmdArgstr & "Y" & .DeviceY
    
        'Command  D
        If .SlideINOUT <> -1 Then cmdArgstr = cmdArgstr & "D" & .SlideINOUT
    
    End With

    Get_IlluminatorParameter_IA = cmdArgstr

    If cmdArgstr = vbNullString Then
        MsgBox "Illegal Command Get_IlluminatorParameter_IA[" & strIllumMode & "]" & vbCrLf & "ErrCode.4-5-10-5-50", vbExclamation, "<<Error>>@Get_IlluminatorParameter_IA"
        Get_IlluminatorParameter_IA = -1
    End If
            
End Function
'==================================
'========== InterAction ===========
'==================================
Private Function CheckGpibStatusFlg() As Boolean

    Dim wkshtObj As Object
    Dim basePoint As Variant
    
    '======= WorkSheet Select ========
    Set wkshtObj = ThisWorkbook.Sheets(IA_WRKSHT_NAME)
         
    '======= WorkSheet ErrorProcess ========
    If IsEmpty(wkshtObj) Then
        MsgBox "Non OptSet[" & IA_WRKSHT_NAME & "]WorkSheet For IA" & vbCrLf & "ErrCode.4-5-11-8-51", vbExclamation, "<<Error>>@CheckGpibStatusFlg"
        Exit Function
    End If

    '======= Base Point Find ========
    Set basePoint = Worksheets(IA_WRKSHT_NAME).Range("A1:F10").Find("GPIB Status")
    If basePoint Is Nothing Then
        MsgBox "Search Error! Not Finding GPIB Status KeyWord @[" & IA_WRKSHT_NAME & "]WorkSheet For IA" & vbCrLf & "ErrCode.4-5-11-8-52", vbExclamation, "<<Error>>@CheckGpibStatusFlg"
        Exit Function
    End If
    
    '======= Judgement GPIB Status Check ========
    If wkshtObj.Cells(basePoint.Row, basePoint.Column + 1) = "CHECK" Then
        CheckGpibStatusFlg = True
    Else
        CheckGpibStatusFlg = False
    End If
    
End Function
'>>>2011/8/26 M.IMAMURA imitation From otherFunction
Public Function ReadOptLux_Test(Optional blnSetMax As Boolean = False) As Double

    Dim sndCmd As String
    Dim strOptcom As String
    Dim strOptcomRet As String
    
    If hOptPort = 0 Then
        Call NsisIni
    End If
    
    If blnSetMax = True Then
'        If OptCond.IllumModel = "N-SIS3KAI" Then
'            Call OptMod("PIN")
'        End If
        Call OptSend_GPIBCommand("R")
'        Flg_FUnit = FUnit_Pin
'        Flg_DownPosi = DownPosi_Up
    End If
    
    If Flg_DefaultLB = True And blnSetMax = True Then
        Select Case OptCond.IllumModel
            Case "N-SIS3", "N-SIS3KAI"
                'T-Command Write
                sndCmd = "T2"
                Call ibwrt(hOptPort, sndCmd)
                TheHdw.WAIT 0.02
            
            Case "N-SIS5", "N-SIS5KAI"
                'J-Command Write
                sndCmd = "J5"
                Call ibwrt(hOptPort, sndCmd)
                TheHdw.WAIT 0.02
        
        End Select
    End If

    'K-Command Write
    sndCmd = "K1" + Chr(13) + Chr(10)
    Call ibwrt(hOptPort, sndCmd)
    
    TheHdw.WAIT 0.02
    
    strOptcomRet = "0000000000"
    strOptcom = ""
    
    'K-Command Read
    Call ibrd(hOptPort, strOptcomRet)
    strOptcom = strOptcom + strOptcomRet

    If strOptcom = "0000000000" Then
        Call sub_errPALS("Illuminator ERROR !!" & "LuxReturnValue from Illuminator is Nothing", "4-5-12-9-53", Enm_ErrFileBank_LOCAL)
'        MsgBox "Illuminator ERROR !!" & Chr(13) & Chr(10) & "LuxReturnValue from Illuminator is Nothing" & vbCrLf & "ErrCode.4-5-12-9-53", vbOKOnly + vbCritical, "<<Error>>@ReadOptLux"
        Exit Function
    End If
    
    ReadOptLux_Test = CDbl(Mid(strOptcom, 3, 8))
'    If blnSetMax = False Then
'        Call OptStatus
'    End If

End Function
'<<<2011/8/26 M.IMAMURA imitation From otherFunction

'>>>2011/8/26 M.IMAMURA imitation From otherFunction
Public Function ReadOptLux(Optional blnSetMax As Boolean = False) As Double

    Dim sndCmd As String
    Dim strOptcom As String
    Dim strOptcomRet As String
    
    If hOptPort = 0 Then
        Call NsisIni
    End If
    
    If blnSetMax = True Then
'        If OptCond.IllumModel = "N-SIS3KAI" Then
'            Call OptMod("PIN")
'        End If
        Call OptSend_GPIBCommand("R")
'        Flg_FUnit = FUnit_Pin
'        Flg_DownPosi = DownPosi_Up
    End If
    
    If Flg_DefaultLB = True And blnSetMax = True Then
        Select Case OptCond.IllumModel
            Case "N-SIS3", "N-SIS3KAI"
                'T-Command Write
                sndCmd = "T2"
                Call ibwrt(hOptPort, sndCmd)
                TheHdw.WAIT 0.02
            
            Case "N-SIS5", "N-SIS5KAI"
                'J-Command Write
                sndCmd = "J5"
                Call ibwrt(hOptPort, sndCmd)
                TheHdw.WAIT 0.02
        
        End Select
    End If

    'K-Command Write
    sndCmd = "K1" + Chr(13) + Chr(10)
    Call ibwrt(hOptPort, sndCmd)
    
    TheHdw.WAIT 0.02
    
    strOptcomRet = "0000000000"
    strOptcom = ""
    
    'K-Command Read
    Call ibrd(hOptPort, strOptcomRet)
    strOptcom = strOptcom + strOptcomRet

    If strOptcom = "0000000000" Then
        Call sub_errPALS("Illuminator ERROR !!" & "LuxReturnValue from Illuminator is Nothing", "4-5-12-9-53", Enm_ErrFileBank_LOCAL)
'        MsgBox "Illuminator ERROR !!" & Chr(13) & Chr(10) & "LuxReturnValue from Illuminator is Nothing" & vbCrLf & "ErrCode.4-5-12-9-53", vbOKOnly + vbCritical, "<<Error>>@ReadOptLux"
        Exit Function
    End If
    
    ReadOptLux = CDbl(Mid(strOptcom, 3, 8))
    If blnSetMax = False Then
        Call OptStatus
    End If

End Function
'<<<2011/8/26 M.IMAMURA imitation From otherFunction

Public Function OptStatusCheck(CheckString As String)

    Dim tmpArray() As String

    CheckString = Application.WorksheetFunction.Clean(CheckString)    'delete ESCAPE SEQUENCE
    CheckString = Trim(CheckString)                                   'delete SPACE
    tmpArray() = Split(CheckString, ",")                              'Separate String

    OptStatusCheck = tmpArray
    
End Function

'>>>2011/06/13 M.IMAMURA OptReset Add.
Public Function sub_CheckOptCond()
    If OptCond Is Nothing Then
        Call OptIni
    End If
End Function
'<<<2011/06/13 M.IMAMURA OptReset Add.

'>>>2011/08/26 M.IMAMURA NsisIni Add.
Private Sub NsisIni()
    Call ibdev(0, OptCond.IllumGpibAddr, 0, 13, 1, &H13, hOptPort)
End Sub
'<<<2011/08/26 M.IMAMURA NsisIni Add.

Public Sub OptCheck()

    If Flg_Tenken = 0 Then Exit Sub

    Dim StatusString() As String
    Dim optcom As String
  
    Opt_Lux = LUX_ERROR_VALUE
    OptResult = 1
    optcom = Space(100)
    
    Opt_Lux = ReadOptLux(True)
    
    If Flg_Tenken = 1 Then
       TheExec.Datalog.WriteComment "MAX_LUX= " & Opt_Lux & "Lux"
    End If

    OptResult = 0

    If Flg_DefaultLB = False Then
        Select Case OptCond.IllumModel
            Case "N-SIS3", "N-SIS3KAI"
                If OptLimit_NSIS3 < Opt_Lux Then OptResult = 1
            Case "N-SIS5", "N-SIS5KAI"
                If OptLimit_NSIS5 < Opt_Lux Then OptResult = 1
            Case Else
                If OptLimit_NSIS3 < OptLimit_ERROR Then OptResult = 1
                MsgBox "Please select N-SIS3 or N-SIS5 ,Your IllumModel is " & OptCond.IllumModel & "@OptCheck"
        End Select
    Else
        'LB/F HardWare Check!! FromIMX184
        Select Case OptCond.IllumModel
            Case "N-SIS3", "N-SIS3KAI", "N-SIS5KAI"
                If OptLimit_NSIS3LB_HIGH > Opt_Lux Then OptResult = 1
            Case "N-SIS5"
                If OptLimit_NSIS5LB_LOW < Opt_Lux Then OptResult = 1
            Case Else
                If OptLimit_NSIS3 < OptLimit_ERROR Then OptResult = 1
                MsgBox "Please select N-SIS3 or N-SIS5 ,Your IllumModel is " & OptCond.IllumModel & "@OptCheck"
        End Select
    End If
End Sub

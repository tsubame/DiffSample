Attribute VB_Name = "CUS_Functional_T"
Option Explicit

' IG/XL Functional Test Template
' (c) Teradyne, Inc, 1997, 1998, 1999
'     All Rights Reserved
' Inclusion of a copyright notice does not imply that this software has been
' published.
' This software is the trade secret information of Teradyne, Inc.
' Use of this software is only in accordance with the terms of a license
' agreement from Teradyne, Inc.
'
' Revision History:
' Date        Description
' 09/27/99    Release 3.30 Development
'
'Ver1.1 H.Arikawa 2013/02/01 PreBodyÇ…StoppatternÇí«â¡ÅB

Dim Arg_DcCategory As String, Arg_DcSelector As String, _
Arg_AcCategory As String, Arg_AcSelector As String, _
Arg_Timing As String, Arg_Edgeset As String, _
Arg_Levels As String, Arg_Patterns As String, _
Arg_StartOfBodyF As String, Arg_PrePatF As String, _
Arg_PreTestF As String, Arg_PostTestF As String, _
Arg_PostPatF As String, Arg_EndOfBodyF As String, _
Arg_ReportResult As String, Arg_DriverLO As String, _
Arg_DriverHI  As String, Arg_DriverZ As String, _
Arg_FloatPins As String, Arg_StartOfBodyFInput As String, _
Arg_PrePatFInput As String, Arg_PreTestFInput As String, _
Arg_PostTestFInput As String, Arg_PostPatFInput As String, _
Arg_EndOfBodyFInput As String, Arg_Util1 As String, _
Arg_Util0 As String, Arg_WaitFlags As String, _
Arg_FlagWaitTimeout As String, Arg_PatFlagF As String, _
Arg_PatFlagFInput As String, Arg_RelayMode As String, _
Arg_PatThreading As String, Arg_MatchAllSites As String

Private Const ARGNUM_PATTERNS = 0
Private Const ARGNUM_STARTOFBODYF = 1
Private Const ARGNUM_PREPATF = 2
Private Const ARGNUM_PRETESTF = 3
Private Const ARGNUM_POSTTESTF = 4
Private Const ARGNUM_POSTPATF = 5
Private Const ARGNUM_ENDOFBODYF = 6
Private Const ARGNUM_REPORTRESULT = 7
Private Const ARGNUM_NotUsed = 8
Private Const ARGNUM_DRIVERLO = 9
Private Const ARGNUM_DRIVERHI = 10
Private Const ARGNUM_DRIVERZ = 11
Private Const ARGNUM_FLOATPINS = 12
Private Const ARGNUM_STARTOFBODYFINPUT = 13
Private Const ARGNUM_PREPATFINPUT = 14
Private Const ARGNUM_PRETESTFINPUT = 15
Private Const ARGNUM_POSTTESTFINPUT = 16
Private Const ARGNUM_POSTPATFINPUT = 17
Private Const ARGNUM_ENDOFBODYFINPUT = 18
Private Const ARGNUM_UTIL1 = 19
Private Const ARGNUM_UTIL0 = 20
Private Const ARGNUM_WAITFLAGS = 21
Private Const ARGNUM_FLAGWAITTIMEOUT = 22
Private Const ARGNUM_PATFLAGFNAME = 23
Private Const ARGNUM_PATFLAGFINPUT = 24
Private Const ARGNUM_RELAYMODE = 25
Private Const ARGNUM_PATTHREADING = 26
Private Const ARGNUM_MATCHALLSITES = 27
Private Const ARGNUM_MAXARG = ARGNUM_MATCHALLSITES


' States of driver features which are saved and restored
Private tt_OldPatThreading As Long
Private tt_OldFlagMatchEnable As Boolean
Private tt_OldWaitFlagsTrue As Long
Private tt_OldWaitFlagsFalse As Long
Private tt_OldMatchAllSites As Boolean


' The TestTemplate function simply calls the PreBody, Body, and PostBody
' functions.  The TestTemplate function is called from the tester executive
' code during normal execution rather than calling the PreBody, Body, and
' PostBody individually as a performance optimization.
Function TestTemplate() As Integer
    Dim PreBodyResult As Integer

    ' Call PreBody, the code setting up general timing & levels, registering
    '   functions, and initializing hardware sub-systems
    PreBodyResult = PreBody()

    If PreBodyResult = TL_SUCCESS Then
        ' Call Body, the code performing DUT testing, and also used during
        '   test debug looping
        Call Body

        ' Call PostBody, the code verifying proper test execution, clearing
        '   hw&sw registers as needed
        Call PostBody

        TestTemplate = TL_SUCCESS
    Else
        TestTemplate = TL_ERROR
    End If
End Function

Function PreBody() As Integer
    If TheExec.Flow.IsRunning = False Then Exit Function
    'First, acquire the values of the parameters for this instance
    '   from the Data Manager
    Call GetTemplateParameters
    
    '2013/02/01 LogicDebugReflect
    Call StopPattern
    
    ' Save previous state of pattern threading and set according
    ' to parameter.
    tt_OldPatThreading = TheHdw.Digital.Patgen.Threading
    If Arg_PatThreading = tl_tm_GetIndexOf(TL_C_THREADONSTR) Then
        TheHdw.Digital.Patgen.Threading = True
    Else
        TheHdw.Digital.Patgen.Threading = False
    End If

    ' Register interpose function names with flow control routines which may
    '   need to invoke them
    Call tl_SetInterpose(TL_C_PREPATF, Arg_PrePatF, Arg_PrePatFInput, _
        TL_C_POSTPATF, Arg_PostPatF, Arg_PostPatFInput, _
        TL_C_PRETESTF, Arg_PreTestF, Arg_PreTestFInput, _
        TL_C_POSTTESTF, Arg_PostTestF, Arg_PostTestFInput, _
        TL_C_FLAGMATCHF, Arg_PatFlagF, Arg_PatFlagFInput)

    ' Optionally power down instruments and power supplies
    If (Arg_RelayMode <> TL_C_RELAYPOWERED) Then Call TheHdw.PinLevels.PowerDown

    ' Close Pin-Electronics, High-Voltage, & Power Supply Relays,
    '   of pins noted on the active levels sheet, if needed
''    Call TheHdw.PinLevels.ConnectAllPins  'Delete CUS_Functional_T

    ' Set drive state on specified utility pins.
    If Arg_Util0 <> TL_C_EMPTYSTR Then Call tl_SetUtilState(Arg_Util0, 0)
    If Arg_Util1 <> TL_C_EMPTYSTR Then Call tl_SetUtilState(Arg_Util1, 1)

    ' Instruct functional voltages/currents hardware drivers to acquire
    '   drive/receive values from the DataManager and apply them.
    If (Arg_Levels <> TL_C_EMPTYSTR) Or (Arg_RelayMode <> TL_C_RELAYPOWERED) Then _
        Call TheHdw.PinLevels.ApplyPower

    ' Instruct functional timing hardware drivers to acquire timing values
    '   from the DataManager and apply them.
    If Arg_Timing <> TL_C_EMPTYSTR Then Call TheHdw.Digital.Timing.Load

    ' Set start-state driver conditions on specified pins.
    ' Start state determines the driver value the pin is set to as each pattern
    '   burst starts.
    ' Default is to have start state automatically selected appropriately
    '   depending on the Format of the first vector of each pattern burst.
    If Arg_DriverLO <> TL_C_EMPTYSTR Then _
        Call tl_SetStartState(Arg_DriverLO, chStartLo)
    If Arg_DriverHI <> TL_C_EMPTYSTR Then _
        Call tl_SetStartState(Arg_DriverHI, chStartHi)
    If Arg_DriverZ <> TL_C_EMPTYSTR Then _
        Call tl_SetStartState(Arg_DriverZ, chStartOff)

    ' Set init-state driver conditions on specified pins
    ' Setting init state causes the pin to drive the specified value.  Init
    '   state is set once, during the prebody, before the first pattern burst.
    ' Default is to leave the pin driving whatever value it last drove during
    '   the previous pattern burst.
    If Arg_DriverLO <> TL_C_EMPTYSTR Then _
        Call tl_SetInitState(Arg_DriverLO, chInitLo)
    If Arg_DriverHI <> TL_C_EMPTYSTR Then _
        Call tl_SetInitState(Arg_DriverHI, chInitHi)
    If Arg_DriverZ <> TL_C_EMPTYSTR Then _
        Call tl_SetInitState(Arg_DriverZ, chInitOff)

    ' Initialize the decoded values for the flag condition settings.
    Dim FlagMatchEnable As Boolean
    Dim WaitFlagsTrue As Long
    Dim WaitFlagsFalse As Long
    Dim MatchAllSites As Boolean
    FlagMatchEnable = False
    WaitFlagsTrue = 0
    WaitFlagsFalse = 0
    MatchAllSites = False

    ' Read back state of flag feature for later restoration
    Call TheHdw.Digital.Patgen.GetFlagMatch(tt_OldFlagMatchEnable, tt_OldWaitFlagsTrue, _
                                            tt_OldWaitFlagsFalse, tt_OldMatchAllSites)

    ' Set desired state according to arguments.
    If (InStr(Arg_WaitFlags, "XXXX") = 0) Then
        FlagMatchEnable = True

        ' If needed, decode the 'WaitFlags' string to detect which
        '   of the CPU  pattern wait flags are to be set true, and
        '   which are to be set false.
        Call tl_tm_GetFlagsTrueAndFalse(Arg_WaitFlags, _
                                        WaitFlagsTrue, WaitFlagsFalse)

        If Arg_MatchAllSites = tl_tm_GetIndexOf(TL_C_YNYESSTR) Then
            MatchAllSites = True
        End If
    End If
    Call TheHdw.Digital.Patgen.SetFlagMatch(FlagMatchEnable, WaitFlagsTrue, WaitFlagsFalse, MatchAllSites)

    PreBody = TL_SUCCESS

End Function

Function Body() As Integer
    Dim temp As String
    Dim ReturnStatus As Long
    If TheExec.Flow.IsRunning = False Then Exit Function

    On Error GoTo ErrHandler

    ' Run the 'StartOfBodyF' interpose function, if specified.
    If Arg_StartOfBodyF <> TL_C_EMPTYSTR Then _
        Call TheExec.Flow.CallFuncWithArgs(Arg_StartOfBodyF, Arg_StartOfBodyFInput)

    ' Remove specified DUT pins, if any, from connection to tester
    '   pin-electronics and other resources
    If Arg_FloatPins <> TL_C_EMPTYSTR Then _
        Call tl_SetFloatState(Arg_FloatPins)

    ' enable the timeout counter
    TheHdw.Digital.Patgen.TimeoutEnable = True
    If Arg_FlagWaitTimeout = TL_C_EMPTYSTR Then Arg_FlagWaitTimeout = "0"
    TheHdw.Digital.Patgen.TIMEOUT = CDbl(Arg_FlagWaitTimeout)

    If TheExec.sites.ActiveCount > 0 Then
        Call TheHdw.Digital.Patterns.pat(Arg_Patterns).test( _
              CLng(val(Arg_ReportResult)), CLng(TL_C_YES))
    End If

    ' Run the 'EndOfBodyF' interpose function, if specified
    If Arg_EndOfBodyF <> TL_C_EMPTYSTR Then _
        Call TheExec.Flow.CallFuncWithArgs(Arg_EndOfBodyF, Arg_EndOfBodyFInput)

    Body = TL_SUCCESS
    Exit Function

ErrHandler:
    On Error GoTo 0
    temp = TheExec.DataManager.InstanceName
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & temp)
    Call TheExec.ErrorReport
    Body = TL_ERROR
End Function

Function PostBody() As Integer
    Dim DriverOFF As String
    If TheExec.Flow.IsRunning = False Then Exit Function

    ' Clear previously registered interpose function names
    Call tl_ClearInterpose(TL_C_PREPATF, TL_C_POSTPATF, TL_C_PRETESTF, _
        TL_C_POSTTESTF)

    ' Return channels to the default start-state condition, as needed
    DriverOFF = tl_tm_CombineCslStrings(Arg_DriverHI, Arg_DriverLO)
    If DriverOFF <> TL_C_EMPTYSTR Then Call tl_SetStartState(DriverOFF, chStartOff)

    ' Return specified DUT pins, if any, to connection with tester pin-electronics & power
    If Arg_FloatPins <> TL_C_EMPTYSTR Then Call tl_ConnectTester(Arg_FloatPins)

    ' Restore flag match feature
    Call TheHdw.Digital.Patgen.SetFlagMatch(tt_OldFlagMatchEnable, tt_OldWaitFlagsTrue, _
                                            tt_OldWaitFlagsFalse, tt_OldMatchAllSites)

    ' Restore pattern threading
    TheHdw.Digital.Patgen.Threading = tt_OldPatThreading

    ' Clear the pins that are masked for this test instance
    Call tl_ClearChannelsMaskedTestInstance

    PostBody = TL_SUCCESS
End Function

Sub GetTemplateParameters()
    Dim ArgStr() As String
    Call tl_tm_GetInstanceValues(ARGNUM_MAXARG, ArgStr)

    Arg_DcCategory = ArgStr(TL_C_DCCATCOLNUM)
    Arg_DcSelector = ArgStr(TL_C_DCSELCOLNUM)
    Arg_AcCategory = ArgStr(TL_C_ACCATCOLNUM)
    Arg_AcSelector = ArgStr(TL_C_ACSELCOLNUM)
    Arg_Timing = ArgStr(TL_C_TIMESETCOLNUM)
    Arg_Edgeset = ArgStr(TL_C_EDGESETCOLNUM)
    Arg_Levels = ArgStr(TL_C_LEVELSCOLNUM)

    Arg_Patterns = ArgStr(ARGNUM_PATTERNS)
    Arg_StartOfBodyF = ArgStr(ARGNUM_STARTOFBODYF)
    Arg_PrePatF = ArgStr(ARGNUM_PREPATF)
    Arg_PreTestF = ArgStr(ARGNUM_PRETESTF)
    Arg_PostTestF = ArgStr(ARGNUM_POSTTESTF)
    Arg_PostPatF = ArgStr(ARGNUM_POSTPATF)
    Arg_EndOfBodyF = ArgStr(ARGNUM_ENDOFBODYF)
    Arg_ReportResult = ArgStr(ARGNUM_REPORTRESULT)
    Arg_DriverLO = ArgStr(ARGNUM_DRIVERLO)
    Arg_DriverHI = ArgStr(ARGNUM_DRIVERHI)
    Arg_DriverZ = ArgStr(ARGNUM_DRIVERZ)
    Arg_FloatPins = ArgStr(ARGNUM_FLOATPINS)
    Arg_StartOfBodyFInput = ArgStr(ARGNUM_STARTOFBODYFINPUT)
    Arg_PrePatFInput = ArgStr(ARGNUM_PREPATFINPUT)
    Arg_PreTestFInput = ArgStr(ARGNUM_PRETESTFINPUT)
    Arg_PostTestFInput = ArgStr(ARGNUM_POSTTESTFINPUT)
    Arg_PostPatFInput = ArgStr(ARGNUM_POSTPATFINPUT)
    Arg_EndOfBodyFInput = ArgStr(ARGNUM_ENDOFBODYFINPUT)
    Arg_Util1 = ArgStr(ARGNUM_UTIL1)
    Arg_Util0 = ArgStr(ARGNUM_UTIL0)
    Arg_WaitFlags = ArgStr(ARGNUM_WAITFLAGS)
    Arg_FlagWaitTimeout = ArgStr(ARGNUM_FLAGWAITTIMEOUT)
    Arg_PatFlagF = ArgStr(ARGNUM_PATFLAGFNAME)
    Arg_PatFlagFInput = ArgStr(ARGNUM_PATFLAGFINPUT)
    Arg_RelayMode = ArgStr(ARGNUM_RELAYMODE)
    Arg_PatThreading = ArgStr(ARGNUM_PATTHREADING)
    Arg_MatchAllSites = ArgStr(ARGNUM_MATCHALLSITES)
End Sub


Function DatalogType() As Integer
    DatalogType = logFunctional
End Function

' End of Execution Section

Public Function RunIE(Optional FocusArg As Integer) As Boolean
    tl_tm_FocusArg = FocusArg
    Call tl_fs_ResetIECtrl(tl_tm_InstanceEditor)
    With tl_tm_InstanceEditor
        .Name = "Functional_T"
        .FuncPage = True
        .PatFuncPage = True
        .LevTimPage = True
        .PinPage = True
        .InterposePage = True
        .Caption = TL_C_IEFUNCSTR
        .HelpValue = TL_C_FUNC_HELP
    End With
    'InstanceEditor_IE.Show     'Delete CUS_Functional_T
    Call tl_fs_StartIE          'Add CUS_Functional_T
    'the return value will be true if the 'Apply' button was not enabled and if the workbook was valid when the form initialized
    RunIE = (Not (tl_tm_FormCtrl.ButtonEnabled)) And tl_tm_BookIsValid
End Function


Sub AssignTemplateValues()
    Dim ArgStr() As String
    Call tl_tm_GetInstanceValues(ARGNUM_MAXARG, ArgStr)
    For Each tl_tm_ParThisPar In AllPars
        With tl_tm_ParThisPar
            .ParameterValue = ArgStr(.Argnum)
        End With
    Next

    'if the value is blank, then apply the default value to the spreadsheet and the Arg
    Call tl_tm_ManageDefault(AllPars, ARGNUM_MAXARG)

End Sub
Sub ApplyDefaults()
    Call SetupParameters

    For Each tl_tm_ParThisPar In AllPars
        With tl_tm_ParThisPar
            Call tl_tm_PutDefaultIfNeeded(.Argnum, .defaultvalue)
        End With
    Next
    Call tl_tm_CleanUp

End Sub
Function GetArgNames() As String
    Dim CallSetup As Boolean
    CallSetup = False
    If AllPars.Count = 0 Then
        Call SetupParameters    'acquire the Argument information, if needed
        CallSetup = True
    End If
    GetArgNames = tl_tm_ListArgNames(ARGNUM_MAXARG)
    If CallSetup = True Then Call tl_tm_CleanUp
End Function


Sub SetupParameters()
    Call tl_tm_SetupCatSelValidation
    Call tl_tm_SetupTimLevValidation
    Call tl_tm_SetupOverlayValidation
    Call tl_tm_SetupInterposeValidation(ARGNUM_STARTOFBODYF, ARGNUM_PREPATF, _
        ARGNUM_PRETESTF, ARGNUM_POSTTESTF, ARGNUM_POSTPATF, ARGNUM_ENDOFBODYF)
    Call tl_tm_SetupInterposeInputValidation(ARGNUM_STARTOFBODYFINPUT, ARGNUM_PREPATFINPUT, _
        ARGNUM_PRETESTFINPUT, ARGNUM_POSTTESTFINPUT, ARGNUM_POSTPATFINPUT, ARGNUM_ENDOFBODYFINPUT)
    Call tl_tm_SetupConditioningPinlistsValidation(ARGNUM_DRIVERLO, ARGNUM_DRIVERHI, _
        ARGNUM_DRIVERZ, ARGNUM_FLOATPINS, ARGNUM_UTIL1, ARGNUM_UTIL0)

    'Patterns,
    With tl_tm_ParFuncPatternName
        .AllParsAdd
        .Argnum = ARGNUM_PATTERNS
        .ParameterName = TL_C_PatternsStr
        .tl_tm_ParSetParam
'        .TestIsPat = True
        .ValueChoices = JobData.AllPatNames
        .TestNotBlank = True
    End With
    'SetPassFail,
    With tl_tm_ParFuncSetPassFail
        .AllParsAdd
        .Argnum = ARGNUM_REPORTRESULT
        .ParameterName = TL_C_PassFailRegisterStr
        .tl_tm_ParSetParam
        .ValueChoices = TL_C_SPFALLSTR
        .TestIsLegalChoice = True
        .defaultvalue = tl_tm_GetIndexOf(TL_C_SPFLOGALLSTR)
    End With
    'WaitFlags,
    With tl_tm_ParPatFlags
        .AllParsAdd
        .Argnum = ARGNUM_WAITFLAGS
        .ParameterName = TL_C_WaitFlagsStr
        .tl_tm_ParSetParam
        .defaultvalue = "XXXX"
    End With
    'FlagWaitTimeout,
    With tl_tm_ParPatFlagWaitTime
        .AllParsAdd
        .Argnum = ARGNUM_FLAGWAITTIMEOUT
        .ParameterName = TL_C_FlagTimeOutStr
        .tl_tm_ParSetParam
        .TestPositive = True
        .defaultvalue = "30"
    End With
    'tl_tm_ParPatFlagF,
    With tl_tm_ParPatFlagF
        .AllParsAdd
        .Argnum = ARGNUM_PATFLAGFNAME
        .ParameterName = TL_C_PatFlagFStr
        .tl_tm_ParSetParam
    End With
    'tl_tm_ParPatFlagFInput,
    With tl_tm_ParPatFlagFInput
        .AllParsAdd
        .Argnum = ARGNUM_PATFLAGFINPUT
        .ParameterName = TL_C_StartOfBodyFStr & TL_C_IpfInputStr
        .tl_tm_ParSetParam
    End With
    'RelayMode,
    With tl_tm_ParFuncRelayMode
        .AllParsAdd
        .Argnum = ARGNUM_RELAYMODE
        .ParameterName = TL_C_RelayModeStr
        .tl_tm_ParSetParam
        'the valid choices can change, based upon whether Levels is set non-blank
        .ValueChoices = TL_C_RMALLSTR
        .TestNotBlank = True
        .TestIsLegalChoice = True
        .defaultvalue = tl_tm_GetIndexOf(TL_C_RMHOTSTR)
        Call .SetEnabler(tl_tm_ParPpmuRelayMode, TL_C_NOTBLANK)
    End With

    'Threading,
    With tl_tm_ParFuncThreading
        .AllParsAdd
        .Argnum = ARGNUM_PATTHREADING
        .ParameterName = "Threading"
        .tl_tm_ParSetParam
        .ValueChoices = TL_C_THREADALLSTR
        .TestNotBlank = True
        .TestIsLegalChoice = True
        .defaultvalue = tl_tm_GetIndexOf(TL_C_THREADOFFSTR)
    End With
    'MatchAllSites,
    With tl_tm_ParPatMatchAllSites
        .AllParsAdd
        .Argnum = ARGNUM_MATCHALLSITES
        .ParameterName = "MatchAllSites"
        .tl_tm_ParSetParam
        .ValueChoices = TL_C_YNALLSTR
        .TestNotBlank = True
        .TestIsLegalChoice = True
        .defaultvalue = tl_tm_GetIndexOf(TL_C_YNNEGSTR)
    End With



End Sub

Function ValidateParameters(Optional VDCint As Integer) As Integer
    'This function is used, at validation time, to determine whether the data
    '   to be executed is proper, valid, and copacetic.  It can be called by
    '   an Instance Editor, or by the Job Validation routines.
    Dim TestResult As Integer
    Dim temp As String
    '   This has modes to run in.  If a mode of '0' is specified for
    '   input, it is assumed that the mode is TL_C_VALDATAMODEJOBVAL.
    '   The modes that .ValidateParameters can operate in are:
    '   TL_C_VALDATAMODEJOBVAL  -   Job Validation mode; report errors to sheet.
    '   TL_C_VALDATAMODENORMAL  -   Instance Editor mode; Fix the current parameter being evaluated.
    '   TL_C_VALDATAMODENOSTOP  -   Instance Editor mode; Do not stop to fix any parameters.
    '   It can return with different modes, such as:
    '   TL_C_VALDATAMODENOFIX   -   Instance Editor mode; Error found, that specific one was not fixed.
    '   TL_C_VALDATAMODEFIXNONE -   Instance Editor mode; Error(s) found, none were fixed.

    'Success is first assumed; if a problem is noted, ValidateParameters will be
    '   set to failure by this routine.
    ValidateParameters = TL_SUCCESS

    If VDCint = 0 Then VDCint = TL_C_VALDATAMODEJOBVAL
    If (VDCint <> TL_C_VALDATAMODENORMAL) And (VDCint <> TL_C_VALDATAMODENOSTOP) _
        And (VDCint <> TL_C_VALDATAMODEJOBVAL) Then
        'denote an error
        temp = TheExec.DataManager.InstanceName
        Call TheExec.ErrorLogMessage("ValidateParameters: Improper mode, instance: " & temp)
        Call TheExec.ErrorReport
        ValidateParameters = TL_ERROR
        Exit Function
    End If

    If VDCint = TL_C_VALDATAMODEJOBVAL Then
        With JobData
            'Get list of pins and pin-groups from datatools.
            Call tl_fs_TemplateJobDataPinlistStrings(JobData, VDCint)

            'Get lists of Categories, Selectors, Timesets, Edgesets, and Levels
            Call tl_fs_TemplateCatSelStrings(.AvailDcCat, .AvailDcSel, _
                .AvailAcCat, .AvailAcSel, .AvailTimeSetAll, .AvailTimeSetExtended, _
                .AvailEdgeSet, .AvailLevels)
            'Get list of Overlay
            Call tl_fs_TemplateOverlayString(.AvailOverlay)
        End With

        'Define the Parameter types and tests to be performed
        Call SetupParameters

        'Now, acquire the values of the parameters for this Template Instance
        '   from the DataManager and assign them to the TemplateArg structures.
        Call AssignTemplateValues
    End If

    ValidateParameters = TL_SUCCESS

    ' Choose tests to perform
    Call tl_tm_ChooseTests(AllPars, VDCint)

    ' Now run the tests on each Argument
    Call tl_tm_RunTests(AllPars, VDCint, TestResult)
    If TestResult <> TL_SUCCESS Then ValidateParameters = TL_ERROR
    If (TestResult <> TL_SUCCESS) And (VDCint = TL_C_VALDATAMODENORMAL) Then Exit Function

'    Warning: Be aware that interpose functions are not validated

'Do something to validate the PatFlagF and PatFlagFInput

    If VDCint = TL_C_VALDATAMODEJOBVAL Then Call tl_tm_CleanUp
End Function

Function ValidateDriverParameters() As Integer
    Dim RetVal As Long
    ValidateDriverParameters = TL_SUCCESS
    Call SetupParameters
    'Now, acquire the values of the parameters for this Template Instance
    '   from the DataManager and assign them to the TemplateArg structures.
    Call AssignTemplateValues
    ' Now validate the patterns used
    RetVal = ValPatThreading(tl_tm_ParFuncPatternName, tl_tm_ParFuncThreading)
    If RetVal = TL_ERROR Then
        ValidateDriverParameters = TL_ERROR
    End If
    Call tl_tm_CleanUp
End Function

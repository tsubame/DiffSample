Attribute VB_Name = "XEeeAuto_TestConditionMacro"
'�T�v:
'   TestCondition����Ă΂��}�N���W
'
'�ړI:
'   TestCondition�V�[�g����Ă΂��}�N�����`����
'
'�쐬��:
'   2011/12/07 Ver0.1 D.Maruyama    Draft
'   2011/12/15 Ver0.2 D.Maruyama    FW_CallUserMacro�ǉ�
'   2011/12/16 Ver0.3 D.Maruyama    �ȉ��̊֐���Optional�����Ƃ���Wait��ǉ�
'                                   �EFW_set_voltage
'                                   �EFW_PatRun
'                                   �EFW_PatSet
'   2012/01/23 Ver0.4 D.Maruyama    SUB�d�����莞�̃p�����[�^�ݒ�֐���ǉ�
'                                   �EFW_SetSubCurrentParam
'                                   �EGetSubCurrentAverageCount�i�擾�p�j
'                                   �EGetSubCurrentClampCurrent�i�擾�p�j
'                                   �EGetSubCurrentWaitTime�i�擾�p�j
'                                   �摜�L���v�`���̃p�����[�^�ݒ�֐���ǉ�
'                                   �EFW_SetCaptureAverage
'                                   �EGetCaptureAverageCount�i�擾�p�j
'                                   �EGetCaptureAverageMode�i�擾�p�j
'   2012/02/03 Ver0.5 D.Maruyama    SUB�d�����莞��Key����TestInstace����擾����悤�ɕύX
'                                   �摜�L���v�`���̃p�����[�^�ݒ�֐���FrameSkip���ɒǉ����A�֐������l�[��
'                                   �EFW_SetCaptureParam
'                                   �EGetCaptureParamAverageCount�i�擾�p�j
'                                   �EGetCaptureParamAverageMode�i�擾�p�j
'                                   FrameSkip�擾�p�̊֐���ǉ�
'                                   �EGetCaptureParamFrameSkip�i�擾�p�j
'   2012/02/09 Ver0.6 D.Maruyama    FW_set_voltage�ɂ��Ĉȉ��̑Ώ����s���ϐ���2����
'                                   �ESUB�؂藣���Ɩ���JobRoute�Ɉڊ�
'                                   �EXCLR�̏�����TestCondition�Ɉڊ�
'   2012/02/09 Ver0.7 D.Maruyama    �ȉ��̊֐��̓C���X�^���X�����L�[���ɕt�^�����`�ɕύX�i����SUB����ɑΉ����邽�߁j
'                                   �EFW_SetSubCurrentParam
'                                   �EGetSubCurrentAverageCount�i�擾�p�j
'                                   �EGetSubCurrentClampCurrent�i�擾�p�j
'                                   �EGetSubCurrentWaitTime�i�擾�p�j
'   2012/03/06 Ver0.71 D.Maruyama    �G���[�n���h�����O�����ׂ�OnError�`���ɓ���
'                                   �ȉ��̊֐����폜
'                                   �ESiteLimit
'   2012/03/07 Ver0.8 D.Maruyama    �ȉ��̊֐���ǉ�
'                                   �EFW_WaitSetTopt
'                                     �V�K�쐬��]����FW_WaitSet�
'                                   �EGetScrnMeasureWaitParam
'                                   �EFW_ScrnMeasureWaitParam
'
'                                   �ȉ��̊֐����C��
'                                   �EFW_SET_RELAY_CONDITION
'                                     CUB�̈������Ƃ�悤�ɕύX
'                                   �EFW_DisconnectPins
'                                     �����Z���ɂ܂����ŋL�q�ł���悤�ɕύX
'                                   �EFW_ConnectPins
'                                     FW_ConnectIOPins�ɕύX�PinsConnect�̂�
'                                   �EFW_WaitSet
'                                     �܂�������TheHdw.Wait�ɌŒ�
'                                   �EFW_SetFVMI_BPMU
'                                   �EFW_SetFIMV_BPMU
'                                   �EFW_DisconnectPins_BPMU
'                                     �T�C�g�w����ȗ��\�ɂ�����ȗ����͑S�T�C�g���s
'                                   �EFW_PatternStop
'                                     ���̂�FW_StopPat�ɕύX
'                                   �EConvertStartState
'                                     chStartFmt , chStartNone��ǉ�
'   2012/03/23 Ver0.9 D.Maruyama    �ȉ��̊֐����폜
'                                   �EGetScrnMeasureWaitParam
'                                   �EFW_ScrnMeasureWaitParam
'                                   �ȉ��̊֐���ǉ�
'                                   �EFW_SetScrnMeasureParam
'                                   �EGetScrnMeasureWaitTime
'                                   �EGetScrnMeasureAverageCount
'                                   �EFW_SeparateFailSiteGnd
'   2012/04/06 Ver1.0 D.Maruyama    �ȉ��̊֐����C��
'                                   �EFW_SeparateFailSiteGnd
'                                       SUB�̎擾���֐��̐擪�Ɉړ��A�ȗ������̏ꍇ�����ɔ�����悤�ɂ���
'                                   �EFW_set_voltage
'                                       PowerCondition�V�[�g�p�~�ɔ����A�������A�������ƁA�V�[�P���X����2�̈����ɕ���
'   2012/04/09 Ver1.1 D.Maruyama    FW_SeparateFailSiteGnd�֐���FailSite��GND�����[���؂藣���悤�ɒǉ�
'   2012/09/27 Ver1.2 H.Arikawa     �ȉ��̊֐���ǉ�
'                                   �EFW_SetGND
'   2012/10/01 Ver1.3 H.Arikawa     �ȉ��̊֐���Stop�������폜
'                                   �EFW_SET_RELAY_CONDITION
'                                   �EFW_OptSet
'                                   �EFW_set_voltage
'                                   �EFW_ConnectIOPins
'                                   �EFW_DisconnectPins
'                                   �EFW_WaitSet
'                                   �EFW_WaitSetTopt
'                                   �EFW_SetFVMI
'                                   �EFW_SetFIMV
'                                   �EFW_SetFVMI_BPMU
'                                   �EFW_SetFIMV_BPMU
'                                   �EFW_DisconnectPins_BPMU
'                                   �EFW_PatSet
'                                   �EFW_PatRun
'                                   �EFW_StopPat
'                                   �EFW_SetIOPinState
'                                   �EFW_SetIOPinElectronics
'                                   �EFW_CallUserMacro
'                                   �EFW_SeparateFailSiteGnd
'                                   �EFW_SetSubCurrentParam
'                                   �EGetSubCurrentAverageCount
'                                   �EGetSubCurrentClampCurrent
'                                   �EGetSubCurrentWaitTime
'                                   �EFW_SetScrnMeasureParam
'                                   �ȉ��̊֐���Stop�������폜���A�e�X�g��~��ǉ�
'                                   FW_SetCaptureParam
'                                   �EGetCaptureParamAverageCount
'                                   �EGetCaptureParamAverageMode
'                                   �EGetCaptureParamFrameSkip
'                                   �EGetScrnMeasureWaitTime
'                                   �EGetScrnMeasureAverageCount
'                                   �EFW_SetGND
'   2012/10/18 Ver1.4 H.Arikawa     �ȉ��̊֐���ǉ�
'                                   �EFW_PatRun_Decoder
'                                   �EFW_PatSet_Decoder
'   2012/10/19 Ver1.5 K.Tokuyoshi   �ȉ��̊֐���ǉ�
'                                   �EGetSubCurrentPinResourceName
'                                   �EFW_SET_RELAY_ON
'                                   �EFW_SET_RELAY_OFF
'                                   �EDutConnectDbNumber
'   2012/10/22 Ver1.8 K.Tokuyoshi   �ȉ��̊֐���ǉ�
'                                   �EFW_SetHoldVoltageParam
'                                   �EGetHoldVoltageAverageCount
'                                   �EGetHoldVoltageClampCurrent
'                                   �EGetHoldVoltageWaitTime
'   2012/10/22 Ver1.9 K.Tokuyoshi   �ȉ��̊֐����C��
'                                   �EFW_SetSubCurrentParam
'   2012/11/02 Ver2.0 H.Arikawa     �ȉ��̊֐���ǉ�
'                                   �EFW_set_voltage_ForUS
'                                   �EPowerDownAndDisconnect
'                                   �ȉ��̊֐����C��
'                                   �EFW_set_voltage
'   2012/12/20 Ver2.1 H.Arikawa     �ȉ��̊֐���ǉ�
'                                   �EFW_PatStatus
'   2013/01/14 Ver2.2 H.Arikawa     �ȉ��̊֐����C��
'                                   �EFW_PatStatus
'                                   �EFW_PatSetTypeSelect
'   2013/01/22 Ver2.3 H.Arikawa     �ȉ��̊֐����C��
'                                   �EFW_SetCaptureParam
'   2013/01/29 Ver2.4 H.Arikawa     �ȉ��̊֐����b��ǉ�
'                                   �EFW_OptEscape
'                                   �EFW_OptModOrModZ1
'                                   �EFW_OptModOrModZ2
'   2013/01/31 Ver2.5 H.Arikawa     �ȉ��̊֐����폜(���g�p�̈�)
'                                   �EFW_SeparateFailSiteGnd
'   2013/02/01 Ver2.6 H.Arikawa     �ȉ��̊֐����C���E�ǉ�
'                                   �EFW_SetSubCurrentParam
'                                   �EFW_SetSubCurrentParam_BPMU
'   2013/02/07 Ver2.7 H.Arikawa     �ȉ��̊֐����C��
'                                   �EFW_SetCaptureParam
'   2013/02/08 Ver2.8 H.Arikawa     �ȉ��̊֐����C��
'                                   �EFW_OptEscape
'                                   �EFW_OptModOrModZ1
'                                   �EFW_OptModOrModZ2
'   2013/02/12 Ver2.9 H.Arikawa     �ȉ��̊֐����C��
'                                   �EFW_SetFIMV
'   2013/02/19 Ver3.0 H.Arikawa     �ȉ��̊֐����C��(�b��)
'                                   �EFW_OptEscape
'                                   �EFW_OptModOrModZ1
'                                   �EFW_OptModOrModZ2
'   2013/02/22 Ver3.1 H.Arikawa     �ȉ��̊֐����C��
'                                   �EFW_ScreeningWait
'   2013/02/22 Ver3.2 H.Arikawa     �ȉ��̊֐���Flg_Simulator�̏�����ǉ�
'                                   �EFW_PatSet
'                                   �EFW_PatSet_Decoder
'                                   �EFW_PatRun
'                                   �EFW_PatRun_Decoder
'                                   �EFW_StopPat
'                                   �EFW_PatSetTypeSelect
'                                   �EFW_PatStatus
'   2013/02/25 Ver3.3 H.Arikawa     �ȉ��̊֐��̊m��ł�ǉ�
'                                   �EFW_OptEscape
'                                   �EFW_OptModOrModZ1
'                                   �EFW_OptModOrModZ2
'   2013/02/28 Ver3.4 H.Arikawa     �ȉ��̊֐����C��
'                                   �EFW_PatSet
'   2013/03/04 Ver3.5 H.Arikawa     �ȉ��̊֐����C��
'                                   �EFW_set_voltage
'                                   �ȉ��̊֐����폜
'                                   �EFW_set_voltage_ForUS
'   2013/03/04 Ver3.6 H.Arikawa     �ȉ��̊֐����C��
'                                   �EFW_set_voltage
'   2013/03/11 Ver3.7 H.Arikawa     �ȉ��̊֐����ȗ���
'                                   �EFW_OptEscape
'                                   �EFW_OptModOrModZ1
'                                   �EFW_OptModOrModZ2
'   2013/03/11 Ver3.8 H.Arikawa     �ȉ��̊֐���ǉ�
'                                   �EFW_DebugWait
'   2013/03/15 Ver3.9 H.Arikawa     �ȉ��̊֐���ǉ�
'                                   �EPatSet
'   2013/08/21 Ver4.0 H.Arikawa     �ȉ��̊֐���ǉ�
'                                   �EFW_PatSetCustomMacroA
'   2013/08/23 Ver4.1 H.Arikawa     �ȉ��̊֐��̌���SKIP FLAG ON����Skip�����ǉ�
'                                   �EFW_OptEscape
'                                   �EFW_OptModOrModZ1
'                                   �EFW_OptModOrModZ2
'   2013/09/27 Ver4.2 H.Arikawa     FW_SetSubCurrentParam��FW_SetSubCurrentParam_BPMU�𓝍�
'                                   �EFW_SetSubCurrentParam
'   2013/10/28 Ver4.3 H.Arikawa     �����ݒ�ȗ��̃t���O��
'   2013/11/05 Ver4.4 T.Morimoto    FW_DcTopt_Set��FW_DcTopt_Measure��ǉ�


Option Explicit

'VarBank�����p
Private Const SUBCURRENT_AVERAGE_COUNT As String = "_SUBCURRENT_AVERAGE_COUNT__"
Private Const SUBCURRENT_CLAMP_CURRENT As String = "_SUBCURRENT_CLAMP_CURRENT__"
Private Const SUBCURRENT_WAIT_TIME As String = "_SUBCURRENT_WAIT_TIME__"
Private Const SUBCURRENT_PIN_RESOURCE As String = "_SUBCURRENT_PIN_RESOURCE__"

Private Const HOLDVOLTAGE_AVERAGE_COUNT As String = "_HOLDVOLTAGE_AVERAGE_COUNT__"
Private Const HOLDVOLTAGE_CLAMP_CURRENT As String = "_HOLDVOLTAGE_CLAMP_CURRENT__"
Private Const HOLDVOLTAGE_WAIT_TIME As String = "_HOLDVOLTAGE_WAIT_TIME__"

Private Const SCRN_MEAS_WAIT_TIME As String = "_SCRNMEAS_WAIT_TIME__"
Private Const SCRN_MEAS_AVERAGE_COUNT As String = "_SCRNMEAS_AVERAGE_COUNT__"

Private Const CAPTURE_PARAM_AVERAGE_COUNT As String = "_CAPTURE_PARAM_AVERAGE_COUNT__"
Private Const CAPTURE_PARAM_AVERAGE_MODE As String = "_CAPTURE_PARAM_AVERAGE_MODE__"
Private Const CAPTURE_PARAM_FRAME_SKIP As String = "_CAPTURE_PARAM_FRAME_SKIP__"

Private Const CAPTURE_AVERAGE_MODE_AVERAGE As String = "Average"
Private Const CAPTURE_AVERAGE_MODE_NO_AVERAGE As String = "NonAverage"

Private OptCheckCounter As Double
Public PatCheckCounter As Double

'���e:
'   �����[�ݒ���s��
'
'�p�����[�^:
'    [Arg0]      In   �����[�Z�b�g��
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_SET_RELAY_CONDITION(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler

    If Parameter.ArgParameterCount() <> 3 Then
        Err.Raise 9999, "FW_SET_RELAY_CONDITION", "The number of FW_SET_RELAY_CONDITION's arguments is invalid." & " @ " & Parameter.ConditionName
    End If

'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_APMU_UB
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================

    '�����[�ݒ�
    Call SET_RELAY_CONDITION(Parameter.Arg(0), Parameter.Arg(1))
    
    If Parameter.Arg(2) <> "-" And Parameter.Arg(2) <> "" Then
        Application.Run Parameter.Arg(2)
    End If
    
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================

    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �����ݒ���s��
'
'�p�����[�^:
'    [Arg0]      In   �����Z�b�g��
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_OptSet(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    OptCheckCounter = 0
    
    If Parameter.ArgParameterCount() < 1 Or Parameter.ArgParameterCount > 2 Then
        Err.Raise 9999, "FW_OptSet", "The number of FW_OptSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------
    
    'Call "OptIni" in case the following object "OptCond" is nothing. (Optcond�I�u�W�F�N�g��Nothing��������OptIni��������)
    Call sub_CheckOptCond
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_ILLUMINATOR
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================
    
    '�����ݒ�
    Call OptSet(Parameter.Arg(1), Parameter.Arg(0))
    
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

Public Sub FW_OptSetAxis(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    OptCheckCounter = 0
    
    If Parameter.ArgParameterCount() < 1 Or Parameter.ArgParameterCount > 2 Then
        Err.Raise 9999, "FW_OptSet", "The number of FW_OptSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------
    
    'Call "OptIni" in case the following object "OptCond" is nothing. (Optcond�I�u�W�F�N�g��Nothing��������OptIni��������)
    Call sub_CheckOptCond
    
''=========Before TestCondition Check Start ======================
'    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
'        Dim eMode As eTestCnditionCheck
'        eMode = TCC_ILLUMINATOR
'        Call CheckBeforeTestCondition(eMode, Parameter)
'    End If
''=========Before TestCondition Check End ========================
    
    '�����ݒ�
    Call OptSet_Axis(Parameter.Arg(1), Parameter.Arg(0))
    
''=========After TestCondition Check Start ======================
'    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
'        Call CheckAfterTestCondition(eMode, Parameter)
'    End If
''=========After TestCondition Check End ========================
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

Public Sub FW_OptSetDevice(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    OptCheckCounter = 0
    
    If Parameter.ArgParameterCount() < 1 Or Parameter.ArgParameterCount > 2 Then
        Err.Raise 9999, "FW_OptSet", "The number of FW_OptSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------
    
    'Call "OptIni" in case the following object "OptCond" is nothing. (Optcond�I�u�W�F�N�g��Nothing��������OptIni��������)
    Call sub_CheckOptCond
    
''=========Before TestCondition Check Start ======================
'    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
'        Dim eMode As eTestCnditionCheck
'        eMode = TCC_ILLUMINATOR
'        Call CheckBeforeTestCondition(eMode, Parameter)
'    End If
''=========Before TestCondition Check End ========================
    
    '�����ݒ�
    Call OptSet_Device(Parameter.Arg(1), Parameter.Arg(0))
    
''=========After TestCondition Check Start ======================
'    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
'        Call CheckAfterTestCondition(eMode, Parameter)
'    End If
''=========After TestCondition Check End ========================
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

Public Sub FW_OptSet_Test(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    OptCheckCounter = 0
    
    If Parameter.ArgParameterCount() < 1 Or Parameter.ArgParameterCount > 2 Then
        Err.Raise 9999, "FW_OptSet", "The number of FW_OptSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------
    
    'Call "OptIni" in case the following object "OptCond" is nothing. (Optcond�I�u�W�F�N�g��Nothing��������OptIni��������)
    Call sub_CheckOptCond
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_ILLUMINATOR
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================
    
    '�����ݒ�
    Call OptSet_Test(Parameter.Arg(1), Parameter.Arg(0))
    
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

Public Sub FW_OptMod(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 5 Then
        Err.Raise 9999, "FW_OptMod", "The number of FW_OptMod's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    OptCheckCounter = 0
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------

    '�����ݒ�
    With Parameter
        Call OptMod(.Arg(0), .Arg(1))
    End With

    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

Public Sub FW_OptJudgement(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
'    If Parameter.ArgParameterCount() <> 5 Then
'        Err.Raise 9999, "FW_OptMod", "The number of FW_OptMod's arguments is invalid." & " @ " & Parameter.ConditionName
'    End If
'
'    OptCheckCounter = 0

    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------

    Call Opt_Judgment_Test(Parameter.Arg(1)) 'For CIS
'    '�����ݒ�
'    With Parameter
'        Call OptMod(.Arg(0), .Arg(1))
'    End With

    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

Public Sub FW_OptStatus(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
''    Exit Sub
    
    If Parameter.ArgParameterCount() <> 1 Then
        Err.Raise 9999, "FW_OptStatus", "The number of FW_OptMod's arguments is invalid." & " @ " & Parameter.ConditionName
    End If

    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------

    '�����ݒ�
    Dim iStatus As Long
    If OptCond.IllumMaker = NIKON Then
        iStatus = NSIS_II.status
        
        While (iStatus <> 0)
            If OptCheckCounter < 999 Then
                TheHdw.TOPT.Recall
                OptCheckCounter = OptCheckCounter + 1
                Call WaitSet(10 * mS)
                Exit Sub
            End If
            iStatus = NSIS_II.status
        Wend
    End If
'    With Parameter
'        Call OptStatus(.Arg(0))
'    End With

    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

Public Sub FW_OptModZ(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 5 Then
        Err.Raise 9999, "FW_OptModZ", "The number of FW_OptMod's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    OptCheckCounter = 0
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------

    '�����ݒ�
    With Parameter
        Call OptModZ_NSIS5(.Arg(0), .Arg(1))
    End With

    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub
'���e:
'   �d���ݒ���s��
'
'�p�����[�^:
'    [Arg0]      In   ������ PowerSuppluyVoltage�V�[�g�ł̖���
'    [Arg1]      In   �V�[�P���X���@PowerSequence�V�[�g�ł̖���
'    [Arg2]      In   ���s��̃E�F�C�g�^�C��(�ȗ��\ �ȗ���Wait�Ȃ�)
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_set_voltage(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 2 And Parameter.ArgParameterCount() <> 3 Then
        Err.Raise 9999, "FW_set_voltage", "The number of FW_set_voltage's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    'CSetFunctionInfo����p�����[�^�̎擾
    Dim strPowerVoltageName As String
    Dim strPowerSequenceName As String
    strPowerVoltageName = Parameter.Arg(0)
    strPowerSequenceName = Parameter.Arg(1)
    
    '�d�������p��PALS�̕ϐ��Ɋi�[�B
    Now_Mode = strPowerVoltageName
         
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_SETVOLTAGE
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================
    
    '�p�^����~
    Call StopPattern 'EeeJob�֐�
    
    Const Zero_V_Con As String = "ZERO"
    Const Zero_V_Con2 As String = "ZERO_V"
    
    If UCase(strPowerVoltageName) = Zero_V_Con Or UCase(strPowerVoltageName) = Zero_V_Con2 Then
        '�d��Condition�̓K�p(For ZERO_V) APMU Pin�ɂ��ẮA5mA�N�����v�A50mA�����W�ɐݒ肷��B
        Call ApplyPowerConditionForUS(strPowerVoltageName, strPowerSequenceName)
    Else:
        '�d��Condition�̓K�p
        Call ApplyPowerCondition(strPowerVoltageName, strPowerSequenceName)
    End If

   '�҂����Ԃ̎w�肪����ꍇ�A�҂�
    If Parameter.ArgParameterCount() = 3 Then
        Dim dblWaitTime As Double
        If Not IsNumeric(Parameter.Arg(2)) Then
            Err.Raise 9999, "FW_set_voltage", "FW_set_voltage's Arg2 is invalid type." & " @ " & Parameter.ConditionName
        End If
        dblWaitTime = Parameter.Arg(2)
        Call TheHdw.WAIT(dblWaitTime)
    End If
    
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================

    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�

End Sub
'���e:
'   �s���̐ڑ����s��
'
'�p�����[�^:
'    [Arg0]      In   �s�����i�s���O���[�v���j
'
'�߂�l:
'
'���ӎ���:
'       ActiveSite���ׂĂ����s�A�T�C�g�V�F�A�֎~
'
Public Sub FW_ConnectIOPins(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 1 Then
        Err.Raise 9999, "FW_ConnectIOPins", "The number of FW_ConnectIOPins's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_TESTER_CHANNELS
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================
    
    Call TheHdw.Digital.relays.Pins(Parameter.Arg(0)).ConnectPins
    
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �s���̊J�����s��
'
'�p�����[�^:
'    [Arg0]      In   �s�����i�s���O���[�v���j
'    [ArgN-1]
'    [ArgN-1]    In   �T�C�g�ԍ�(�ȗ����ꂽ�ꍇ�͑S�T�C�g)
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_DisconnectPins(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount < 1 Then
        Err.Raise 9999, "FW_DisconnectPins", "The number of FW_DisconnectPins's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
        
    Dim strPins As String
    Dim lSite As Long
        
    Dim i As Long
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_TESTER_CHANNELS
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================
    
    strPins = Parameter.Arg(0)
    If Parameter.ArgParameterCount = 1 Then
    
        Call DisconnectPins(strPins)
        
    ElseIf Parameter.ArgParameterCount >= 2 Then
    
        If Parameter.ArgParameterCount > 2 Then
            For i = 1 To Parameter.ArgParameterCount - 2
                strPins = strPins & "," & Parameter.Arg(i)
            Next i
        End If
        
        With Parameter
            If IsNumeric(.Arg(.ArgParameterCount - 1)) Then
                lSite = .Arg(.ArgParameterCount - 1)
                 Call DisconnectPins(strPins, lSite)
            Else
                strPins = strPins & "," & Parameter.Arg(.ArgParameterCount - 1)
                Call DisconnectPins(strPins)
            End If
        End With
        
    End If
    
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================
                    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub


'���e:
'   �w�莞�ԕ�Wait������
'
'�p�����[�^:
'    [Arg0]      In   Wait����(s)
'
'�߂�l:
'
'���ӎ���:
'     TheHdw.Wait�Ŗⓚ���p�ɑ҂�
'
Public Sub FW_WaitSet(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 1 Then
        Err.Raise 9999, "FW_WaitSet", "The number of FW_WaitSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_WaitSet", "FW_WaitSet's Arg0 is invalid type." & " @ " & Parameter.ConditionName
    End If
    
    Dim dblWaitTime As Double
    dblWaitTime = Parameter.Arg(0)
    
    'Wait����
    Call TheHdw.WAIT(dblWaitTime)
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �w�莞�ԕ�Wait������
'
'�p�����[�^:
'    [Arg0]      In   Wait����(s)
'
'�߂�l:
'
'���ӎ���:
'     TheExec.RunOptions.AutoAcquire
'     �ɉ�����Wait�𕪂���A�Ăяo������TOPT���s���Ă��邩�ӎ�����K�v������
'     TOPT���s���łȂ���Ί��҂�������͂��Ȃ�
'
Public Sub FW_WaitSetTopt(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 1 Then
        Err.Raise 9999, "FW_WaitSet", "The number of FW_WaitSetTopt's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_WaitSet", "FW_WaitSetTopt's Arg0 is invalid type." & " @ " & Parameter.ConditionName
    End If
    
    Dim dblWaitTime As Double
    dblWaitTime = Parameter.Arg(0)
    
    'Wait����
    If TheExec.RunOptions.AutoAcquire = True Then
        Call TheHdw.TOPT.WAIT(toptTimer, dblWaitTime * 1000)
    Else
        Call TheHdw.WAIT(dblWaitTime)
    End If
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub


'���e:
'   �s����FVMI�̐ݒ���s��
'
'�p�����[�^:
'    [Arg0]      In   �s�����i�s���O���[�v���j
'    [Arg1]      In   �t�H�[�X�d��
'    [Arg2]      In   �N�����v�d��
'    [Arg3]      In   �T�C�g�ԍ��i-1�͏ȗ��Ƃ��Ĉ����j
'    [Arg4]      In   �R�l�N�g���邩�ǂ���(�ȗ��\�F�ȗ����R�l�N�g)
'                       �iFalse�ŃR�l�N�g���Ȃ��A����ȊO�̓R�l�N�g�j
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_SetFVMI(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 3 And Parameter.ArgParameterCount <> 4 _
        And Parameter.ArgParameterCount <> 5 Then
            Err.Raise 9999, "FW_SetFVMI", "The number of FW_SetFVMI's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPins As String
    Dim dblForce As Double
    Dim dblClamp As Double
    Dim lSite As Long
    Dim strIsConnect As String
    
    strPins = Parameter.Arg(0)
    
    If Not IsNumeric(Parameter.Arg(1)) Then
        Err.Raise 9999, "FW_SetFVMI", "FW_SetFVMI 's Arg1 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblForce = Parameter.Arg(1)
    
    If Not IsNumeric(Parameter.Arg(2)) Then
        Err.Raise 9999, "FW_SetFVMI", "FW_SetFVMI 's Arg2 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblClamp = Parameter.Arg(2)
    
    If Parameter.ArgParameterCount > 3 Then
        If Not IsNumeric(Parameter.Arg(3)) Then
            Err.Raise 9999, "FW_SetFVMI", "FW_SetFVMI 's Arg3 is invalid type." & " @ " & Parameter.ConditionName
        End If
        lSite = Parameter.Arg(3)
    End If
    
    If Parameter.ArgParameterCount = 5 Then
        strIsConnect = Parameter.Arg(4)
    End If
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_TESTER_CHANNELS
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================

    With Parameter
        Select Case .ArgParameterCount
            Case 3
                Call SetFVMI(strPins, dblForce, dblClamp)
            Case 4
                If lSite = -1 Then
                    Call SetFVMI(strPins, dblForce, dblClamp)
                Else
                    Call SetFVMI(strPins, dblForce, dblClamp, lSite)
                End If
            Case 5
                If strIsConnect = "False" Then
                    If lSite = -1 Then
                       Call SetFVMI(strPins, dblForce, dblClamp, , False)
                    Else
                       Call SetFVMI(strPins, dblForce, dblClamp, lSite, False)
                    End If
                Else
                    If lSite = -1 Then
                       Call SetFVMI(strPins, dblForce, dblClamp)
                    Else
                       Call SetFVMI(strPins, dblForce, dblClamp, lSite)
                    End If
                End If
        End Select
    End With
    
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �s����FIMV�̐ݒ���s��
'
'�p�����[�^:
'    [Arg0]      In   �s�����i�s���O���[�v���j
'    [Arg1]      In   �t�H�[�X�d��
'    [Arg2]      In   �N�����v�d��
'    [Arg3]      In   �T�C�g�ԍ��i-1�͏ȗ��Ƃ��Ĉ����j
'    [Arg4]      In   �R�l�N�g���邩�ǂ���(�ȗ��\�F�ȗ����R�l�N�g)
'                       �iFalse�ŃR�l�N�g���Ȃ��A����ȊO�̓R�l�N�g�j
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_SetFIMV(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 3 And Parameter.ArgParameterCount <> 4 _
        And Parameter.ArgParameterCount <> 5 Then
            Err.Raise 9999, "FW_SetFIMV", "The number of FW_SetFIMV's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPins As String
    Dim dblForce As Double
    Dim dblClamp As Double
    Dim lSite As Long
    Dim strIsConnect As String
    
    strPins = Parameter.Arg(0)
    
    If Not IsNumeric(Parameter.Arg(1)) Then
        Err.Raise 9999, "FW_SetFIMV", "FW_SetFIMV 's Arg1 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblForce = Parameter.Arg(1)
    
    If Not IsNumeric(Parameter.Arg(2)) Then
        If UCase(Parameter.Arg(2)) = "NONE" Then
            dblClamp = 5
        Else:
            Err.Raise 9999, "FW_SetFIMV", "FW_SetFIMV 's Arg2 is invalid type." & " @ " & Parameter.ConditionName
        End If
    Else
        dblClamp = Parameter.Arg(2)
    End If
    
    If Parameter.ArgParameterCount > 3 Then
        If Not IsNumeric(Parameter.Arg(3)) Then
            Err.Raise 9999, "FW_SetFIMV", "FW_SetFIMV 's Arg3 is invalid type." & " @ " & Parameter.ConditionName
        End If
        lSite = Parameter.Arg(3)
    End If
    
    If Parameter.ArgParameterCount = 5 Then
        strIsConnect = Parameter.Arg(4)
    End If
    
    With Parameter
        Select Case .ArgParameterCount
            Case 3
                Call SetFIMV(strPins, dblForce, dblClamp)
            Case 4
                If lSite = -1 Then
                    Call SetFIMV(strPins, dblForce, dblClamp)
                Else
                    Call SetFIMV(strPins, dblForce, dblClamp, lSite)
                End If
            Case 5
                If .Arg(4) = "False" Then
                    If lSite = -1 Then
                       Call SetFIMV(strPins, dblForce, dblClamp, , False)
                    Else
                       Call SetFIMV(strPins, dblForce, dblClamp, lSite, False)
                    End If
                Else
                    If lSite = -1 Then
                       Call SetFIMV(strPins, dblForce, dblClamp)
                    Else
                       Call SetFIMV(strPins, dblForce, dblClamp, lSite)
                    End If
                End If
        End Select
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub


'���e:
'   �s����FVMI�̐ݒ���s���iAPMU�j PowerDown��p�@�����W�F50mA�A�N�����v�F5mA�@�Œ�ƂȂ�
'
'�p�����[�^:
'    [Arg0]      In   �s�����i�s���O���[�v���j
'    [Arg1]      In   �t�H�[�X�d��
'    [Arg2]      In   �N�����v�d��
'    [Arg3]      In   �T�C�g�ԍ��i-1�͏ȗ��Ƃ��Ĉ����j
'    [Arg4]      In   �R�l�N�g���邩�ǂ����iFalse�ŃR�l�N�g���Ȃ��A����ȊO�̓R�l�N�g�j
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_SetFVMI_APMUoff(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 3 And Parameter.ArgParameterCount <> 4 _
        And Parameter.ArgParameterCount <> 5 Then
            Err.Raise 9999, "FW_SetFVMI_APMUoff", "The number of FW_SetFVMI_APMUoff's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPins As String
    Dim dblForce As Double
    Dim dblClamp As Double
    Dim lSite As Long
    Dim strIsConnect As String
    
    strPins = Parameter.Arg(0)
    
    If Not IsNumeric(Parameter.Arg(1)) Then
        Err.Raise 9999, "FW_SetFVMI_APMUoff", "FW_SetFVMI_APMUoff 's Arg1 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblForce = Parameter.Arg(1)
    
    If Not IsNumeric(Parameter.Arg(2)) Then
        Err.Raise 9999, "FW_SetFVMI_APMUoff", "FW_SetFVMI_APMUoff 's Arg2 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblClamp = Parameter.Arg(2)
    
    If Parameter.ArgParameterCount > 3 Then
        If Not IsNumeric(Parameter.Arg(3)) Then
            Err.Raise 9999, "FW_SetFVMI_APMUoff", "FW_SetFVMI_APMUoff 's Arg3 is invalid type." & " @ " & Parameter.ConditionName
        End If
        lSite = Parameter.Arg(3)
    End If
    
    If Parameter.ArgParameterCount = 5 Then
        strIsConnect = Parameter.Arg(4)
    End If
    
    With Parameter
        Select Case .ArgParameterCount
            Case 3
                Call SetFVMI_APMUoff(strPins, dblForce, dblClamp)
            Case 4
                If lSite = -1 Then
                    Call SetFVMI_APMUoff(strPins, dblForce, dblClamp)
                Else
                    Call SetFVMI_APMUoff(strPins, dblForce, dblClamp, lSite)
                End If
            Case 5
                If strIsConnect = "False" Then
                    If lSite = -1 Then
                       Call SetFVMI_APMUoff(strPins, dblForce, dblClamp, , False)
                    Else
                       Call SetFVMI_APMUoff(strPins, dblForce, dblClamp, lSite, False)
                    End If
                Else
                    If lSite = -1 Then
                       Call SetFVMI_APMUoff(strPins, dblForce, dblClamp)
                    Else
                       Call SetFVMI_APMUoff(strPins, dblForce, dblClamp, lSite)
                    End If
                End If
        End Select
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub


'���e:
'   �s����FVMI�̐ݒ���s���iBPMU�j
'
'�p�����[�^:
'    [Arg0]      In   �s�����i�s���O���[�v���j
'    [Arg1]      In   �t�H�[�X�d��
'    [Arg2]      In   �N�����v�d��
'    [Arg3]      In   �T�C�g�ԍ��i-1�͏ȗ��Ƃ��Ĉ����j
'    [Arg4]      In   �R�l�N�g���邩�ǂ����iFalse�ŃR�l�N�g���Ȃ��A����ȊO�̓R�l�N�g�j
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_SetFVMI_BPMU(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 3 And Parameter.ArgParameterCount <> 4 _
        And Parameter.ArgParameterCount <> 5 Then
            Err.Raise 9999, "FW_SetFVMI_BPMU", "The number of FW_SetFVMI_BPMU's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPins As String
    Dim dblForce As Double
    Dim dblClamp As Double
    Dim lSite As Long
    Dim strIsConnect As String
    
    strPins = Parameter.Arg(0)
    
    If Not IsNumeric(Parameter.Arg(1)) Then
        Err.Raise 9999, "FW_SetFVMI_BPMU", "FW_SetFVMI_BPMU 's Arg1 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblForce = Parameter.Arg(1)
    
    If Not IsNumeric(Parameter.Arg(2)) Then
        Err.Raise 9999, "FW_SetFVMI_BPMU", "FW_SetFVMI_BPMU 's Arg2 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblClamp = Parameter.Arg(2)
    
    If Parameter.ArgParameterCount > 3 Then
        If Not IsNumeric(Parameter.Arg(3)) Then
            Err.Raise 9999, "FW_SetFVMI_BPMU", "FW_SetFVMI_BPMU 's Arg3 is invalid type." & " @ " & Parameter.ConditionName
        End If
        lSite = Parameter.Arg(3)
    End If
    
    If Parameter.ArgParameterCount = 5 Then
        strIsConnect = Parameter.Arg(4)
    End If
    
    With Parameter
        Select Case .ArgParameterCount
            Case 3
                Call SetFVMI_BPMU(strPins, dblForce, dblClamp)
            Case 4
                If lSite = -1 Then
                    Call SetFVMI_BPMU(strPins, dblForce, dblClamp)
                Else
                    Call SetFVMI_BPMU(strPins, dblForce, dblClamp, lSite)
                End If
            Case 5
                If strIsConnect = "False" Then
                    If lSite = -1 Then
                       Call SetFVMI_BPMU(strPins, dblForce, dblClamp, , False)
                    Else
                       Call SetFVMI_BPMU(strPins, dblForce, dblClamp, lSite, False)
                    End If
                Else
                    If lSite = -1 Then
                       Call SetFVMI_BPMU(strPins, dblForce, dblClamp)
                    Else
                       Call SetFVMI_BPMU(strPins, dblForce, dblClamp, lSite)
                    End If
                End If
        End Select
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �s����FIMV�̐ݒ���s��(BPMU)
'
'�p�����[�^:
'    [Arg0]      In   �s�����i�s���O���[�v���j
'    [Arg1]      In   �t�H�[�X�d��
'    [Arg2]      In   �N�����v�d��
'    [Arg3]      In   �T�C�g�ԍ��i-1�͏ȗ��Ƃ��Ĉ����j
'    [Arg4]      In   �R�l�N�g���邩�ǂ���(�ȗ��\�F�ȗ����R�l�N�g)
'                       �iFalse�ŃR�l�N�g���Ȃ��A����ȊO�̓R�l�N�g�j
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_SetFIMV_BPMU(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 3 And Parameter.ArgParameterCount <> 4 _
        And Parameter.ArgParameterCount <> 5 Then
            Err.Raise 9999, "FW_SetFIMV_BPMU", "The number of FW_SetFIMV_BPMU's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPins As String
    Dim dblForce As Double
    Dim dblClamp As Double
    Dim lSite As Long
    Dim strIsConnect As String
    
    strPins = Parameter.Arg(0)
    
    If Not IsNumeric(Parameter.Arg(1)) Then
        Err.Raise 9999, "FW_SetFIMV_BPMU", "FW_SetFIMV_BPMU 's Arg1 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblForce = Parameter.Arg(1)
    
    If Not IsNumeric(Parameter.Arg(2)) Then
        Err.Raise 9999, "FW_SetFIMV_BPMU", "FW_SetFIMV_BPMU 's Arg2 is invalid type." & " @ " & Parameter.ConditionName
    End If
    dblClamp = Parameter.Arg(2)
    
    If Parameter.ArgParameterCount > 3 Then
        If Not IsNumeric(Parameter.Arg(3)) Then
            Err.Raise 9999, "FW_SetFIMV_BPMU", "FW_SetFIMV_BPMU 's Arg3 is invalid type." & " @ " & Parameter.ConditionName
        End If
        lSite = Parameter.Arg(3)
    End If
    
    If Parameter.ArgParameterCount = 5 Then
        strIsConnect = Parameter.Arg(4)
    End If
    
    With Parameter
        Select Case .ArgParameterCount
            Case 3
                Call SetFIMV_BPMU(strPins, dblForce, dblClamp)
            Case 4
                If lSite = -1 Then
                    Call SetFIMV_BPMU(strPins, dblForce, dblClamp)
                Else
                    Call SetFIMV_BPMU(strPins, dblForce, dblClamp, lSite)
                End If
            Case 5
                If .Arg(4) = "False" Then
                    If lSite = -1 Then
                       Call SetFIMV_BPMU(strPins, dblForce, dblClamp, , False)
                    Else
                       Call SetFIMV_BPMU(strPins, dblForce, dblClamp, lSite, False)
                    End If
                Else
                    If lSite = -1 Then
                       Call SetFIMV_BPMU(strPins, dblForce, dblClamp)
                    Else
                       Call SetFIMV_BPMU(strPins, dblForce, dblClamp, lSite)
                    End If
                End If
        End Select
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �s���̊J�����s��(BPMU)
'
'�p�����[�^:
'    [Arg0]      In   �s�����i�s���O���[�v���j
'    [Arg3]      In   �T�C�g�ԍ��i-1�͏ȗ��Ƃ��Ĉ����j
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_DisconnectPins_BPMU(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 1 And Parameter.ArgParameterCount <> 2 Then
        Err.Raise 9999, "DisconnectPins_BPMU", "The number of DisconnectPins_BPMU's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
        
    Dim strPins As String
    Dim lSite As Long
        
    strPins = Parameter.Arg(0)
    
    If Parameter.ArgParameterCount = 2 Then
        If Not IsNumeric(Parameter.Arg(1)) Then
            Err.Raise 9999, "DisconnectPins_BPMU", "DisconnectPins_BPMU 's Arg1 is invalid type." & " @ " & Parameter.ConditionName
        End If
        lSite = Parameter.Arg(1)
    End If
    
    If Parameter.ArgParameterCount = 1 Then
        Call DisconnectPins(strPins)
    ElseIf Parameter.ArgParameterCount = 2 Then
        Call DisconnectPins(strPins, lSite)
    End If
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �p�^���̊J�n�������Ȃ�(�I�����܂��Ȃ�)
'
'�p�����[�^:
'    [Arg0]      In   �p�^����
'    [Arg1]      In   TSB�V�[�g��
'    [Arg2]      In   ���s��̃E�F�C�g�^�C��(�ȗ��\ �ȗ���Wait�Ȃ�)
'
'�߂�l:
'
'���ӎ���:IP750 or Decoder Pat�́A��p�Őݒ肷��B
'
Public Sub FW_PatSet(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
'    Const PAT_START_LABEL As String = "pat_start"
    Const PAT_START_LABEL As String = "START"

    If Parameter.ArgParameterCount <> 2 And Parameter.ArgParameterCount <> 3 Then
         Err.Raise 9999, "FW_PatSet", "The number of FW_PatSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPatGroupName As String
    Dim strTsbName As String
        
    With Parameter
        strPatGroupName = .Arg(0)
        strTsbName = .Arg(1)
    End With
            
    Call StopPattern_Halt 'EeeJob�֐�
    Call SetTimeOut 'EeeJob�֐�
    
    '�������W�X�^�Ή����[�`�� Start
    '���W�X�^�ݒ蕔Only(keep_alive)�FPatRun
    '���W�X�^�ݒ�+�쓮��:PatSet
    Dim tmpPatGroupName() As String
    Dim i As Integer
    tmpPatGroupName = Split(strPatGroupName, ",")
    
    PatCheckCounter = 0
    
    For i = 0 To UBound(tmpPatGroupName)
        If i < UBound(tmpPatGroupName) Then
            With TheHdw.Digital
                Call .Timing.Load(strTsbName)
                Call .Patterns.pat(tmpPatGroupName(i)).Run(PAT_START_LABEL)
            End With
            If TheExec.RunOptions.AutoAcquire = True Then
                Dim iStatus As Long
                If TheHdw.Digital.Patgen.IsRunningAnySite = True Then      'True:Still Running
                    iStatus = 0
                ElseIf TheHdw.Digital.Patgen.IsRunningAnySite = False Then 'False:haltend or keepalive
                    iStatus = 1
                End If
                
                While (iStatus <> 1)
                    If PatCheckCounter < 999 Then
                        TheHdw.TOPT.Recall
                        PatCheckCounter = PatCheckCounter + 1
                        Call WaitSet(10 * mS)
                        Exit Sub
                    End If
                Wend
            End If
        ElseIf i = UBound(tmpPatGroupName) Then
            With TheHdw.Digital
                Call .Timing.Load(strTsbName)
                Call .Patterns.pat(tmpPatGroupName(i)).Start(PAT_START_LABEL)
                    If Flg_Scrn = 1 And tmpPatGroupName(i) = "PG_CUR_SCR" Then
                        Dim Hsn(nSite) As Double
                        Dim site As Long
                            TheHdw.WAIT 50 * mS
                            Call MeasureI_APMU("P_HSN", Hsn, 50)
                        TheResult.Add "IDDBI_HSN", Hsn
                    End If
            End With
        End If
    Next i
    '�������W�X�^�Ή����[�`�� End
    
   '�҂����Ԃ̎w�肪����ꍇ�A�҂�
    If Parameter.ArgParameterCount() = 3 Then
        Dim dblWaitTime As Double
        If Not IsNumeric(Parameter.Arg(2)) Then
            Err.Raise 9999, "FW_PatSet", "FW_PatSet's Arg2 is invalid type." & " @ " & Parameter.ConditionName
        End If
        dblWaitTime = Parameter.Arg(2)
        Call TheHdw.WAIT(dblWaitTime)
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �p�^���̊J�n�������Ȃ�(�I�����܂��Ȃ�)
'
'�p�����[�^:
'    [Arg0]      In   �p�^����
'    [Arg1]      In   TSB�V�[�g��
'    [Arg2]      In   ���s��̃E�F�C�g�^�C��(�ȗ��\ �ȗ���Wait�Ȃ�)
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_PatSet_Decoder(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
    Const PAT_START_LABEL As String = "pat_start"

    If Parameter.ArgParameterCount <> 2 And Parameter.ArgParameterCount <> 3 Then
         Err.Raise 9999, "FW_PatSet", "The number of FW_PatSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPatGroupName As String
    Dim strTsbName As String
        
    With Parameter
        strPatGroupName = .Arg(0)
        strTsbName = .Arg(1)
    End With
    
    Call StopPattern 'EeeJob�֐�
    Call SetTimeOut 'EeeJob�֐�
    
    With TheHdw.Digital
        Call .Timing.Load(strTsbName)
        Call .Patterns.pat(strPatGroupName).Start(PAT_START_LABEL)
    End With
    
   '�҂����Ԃ̎w�肪����ꍇ�A�҂�
    If Parameter.ArgParameterCount() = 3 Then
        Dim dblWaitTime As Double
        If Not IsNumeric(Parameter.Arg(2)) Then
            Err.Raise 9999, "FW_PatSet", "FW_PatSet's Arg2 is invalid type." & " @ " & Parameter.ConditionName
        End If
        dblWaitTime = Parameter.Arg(2)
        Call TheHdw.WAIT(dblWaitTime)
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �p�^���̊J�n�������Ȃ�(�I�����܂�)
'
'�p�����[�^:
'    [Arg0]      In   �p�^����
'    [Arg1]      In   TSB�V�[�g��
'    [Arg2]      In   ���s��̃E�F�C�g�^�C��(�ȗ��\ �ȗ���Wait�Ȃ�)
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_PatRun(ByVal Parameter As CSetFunctionInfo)

    
    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
'    Const PAT_START_LABEL As String = "pat_start"
    Const PAT_START_LABEL As String = "START"

    If Parameter.ArgParameterCount <> 2 And Parameter.ArgParameterCount <> 3 Then
         Err.Raise 9999, "FW_PatSet", "The number of FW_PatSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPatGroupName As String
    Dim strTsbName As String
        
    With Parameter
        strPatGroupName = .Arg(0)
        strTsbName = .Arg(1)
    End With
    
    Call StopPattern_Halt 'EeeJob�֐�
    Call SetTimeOut 'EeeJob�֐�
    
    With TheHdw.Digital
        Call .Timing.Load(strTsbName)
        Call .Patterns.pat(strPatGroupName).Run(PAT_START_LABEL)
    End With
    
   '�҂����Ԃ̎w�肪����ꍇ�A�҂�
    If Parameter.ArgParameterCount() = 3 Then
        Dim dblWaitTime As Double
        If Not IsNumeric(Parameter.Arg(2)) Then
            Err.Raise 9999, "FW_PatSet", "FW_PatSet's Arg2 is invalid type." & " @ " & Parameter.ConditionName
        End If
        dblWaitTime = Parameter.Arg(2)
        Call TheHdw.WAIT(dblWaitTime)
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �p�^���̊J�n�������Ȃ�(�I�����܂�)
'
'�p�����[�^:
'    [Arg0]      In   �p�^����
'    [Arg1]      In   TSB�V�[�g��
'    [Arg2]      In   ���s��̃E�F�C�g�^�C��(�ȗ��\ �ȗ���Wait�Ȃ�)
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_PatRun_Decoder(ByVal Parameter As CSetFunctionInfo)

    
    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
    Const PAT_START_LABEL As String = "pat_start"

    If Parameter.ArgParameterCount <> 2 And Parameter.ArgParameterCount <> 3 Then
         Err.Raise 9999, "FW_PatSet", "The number of FW_PatSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPatGroupName As String
    Dim strTsbName As String
        
    With Parameter
        strPatGroupName = .Arg(0)
        strTsbName = .Arg(1)
    End With
    
    Call StopPattern 'EeeJob�֐�
    Call SetTimeOut 'EeeJob�֐�
    
    With TheHdw.Digital
        Call .Timing.Load(strTsbName)
        Call .Patterns.pat(strPatGroupName).Run(PAT_START_LABEL)
    End With
    
   '�҂����Ԃ̎w�肪����ꍇ�A�҂�
    If Parameter.ArgParameterCount() = 3 Then
        Dim dblWaitTime As Double
        If Not IsNumeric(Parameter.Arg(2)) Then
            Err.Raise 9999, "FW_PatSet", "FW_PatSet's Arg2 is invalid type." & " @ " & Parameter.ConditionName
        End If
        dblWaitTime = Parameter.Arg(2)
        Call TheHdw.WAIT(dblWaitTime)
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �p�^�����~����
'
'�p�����[�^:
'    �Ȃ�
'
'�߂�l:
'
'���ӎ���:
'   �p�����[�^�͏����Ă����������
'
Public Sub FW_StopPat(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
    Call StopPattern 'EeeJob�֐�
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �s���̏����ݒ���s��
'
'�p�����[�^:
'    [Arg0]      In   �s����
'    [Arg1]      In   InitState[Hi, Lo, Off]
'    [Arg2]      In   StartState[Hi, Lo, Off]
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_SetIOPinState(ByVal Parameter As CSetFunctionInfo)

On Error GoTo ErrHandler

    If Parameter.ArgParameterCount <> 3 Then
         Err.Raise 9999, "FW_SetIOPinState", "The number of FW_SetIOPinState's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPinName As String
    Dim eInitState As ChInitState
    Dim eStartState As chStartState
    
    With Parameter
        strPinName = .Arg(0)
        eInitState = ConvertInitState(.Arg(1))
        eStartState = ConvertStartState(.Arg(2))
    End With
    
    With TheHdw.Pins(strPinName)
        .InitState = eInitState
        .StartState = eStartState
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

Private Function ConvertInitState(ByVal Arg As String) As ChInitState
    
    Select Case Arg
        Case "chInitHi"
            ConvertInitState = chInitHi
        Case "chInitLo"
            ConvertInitState = chInitLo
        Case "chInitOff"
            ConvertInitState = chInitOff
        Case Else
            Err.Raise 9999, "ConvertInitState", "Init State invalide param" '�Ăяo�����ŃG���[�n���h�����O�����Ăق���
    End Select
       
End Function

Private Function ConvertStartState(ByVal Arg As String) As chStartState
    
    Select Case Arg
        Case "chStartHi"
            ConvertStartState = chStartHi
        Case "chStartLo"
            ConvertStartState = chStartLo
        Case "chStartOff"
            ConvertStartState = chStartOff
        Case "chStartFmt"
            ConvertStartState = chStartFmt
        Case "chStartNone"
            ConvertStartState = chStartNone
        Case Else
            Err.Raise 9999, "ConvertStartState", "Start State invalide param" '�Ăяo�����ŃG���[�n���h�����O�����Ăق���
    End Select
        
End Function



'���e:
'   �s���̏����ݒ���s��
'
'�p�����[�^:
'    [Arg0]      In   ClockVoltage�̏�����
'    [Arg1]      In   �Ώۃs���̖��O�iClockVoltage�V�[�g�̖��̂Ɠ��������Ɓj
'
'�߂�l:
'
'���ӎ���:
'       ActiveSite���ׂĂ����s�A�T�C�g�V�F�A�֎~
'
Public Sub FW_SetIOPinElectronics(ByVal Parameter As CSetFunctionInfo)

On Error GoTo ErrHandler

    If Parameter.ArgParameterCount <> 2 Then
         Err.Raise 9999, "FW_SetIOPinElectronics", "The number of FW_SetIOPinElectronics's arguments is invalid." & " @ " & Parameter.ConditionName
    End If

'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_TESTER_CHANNELS
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================
    
    With Parameter
        Call ShtClockV.GetClockInfo(.Arg(0), .Arg(1)).ForceGroupPins(.Arg(1))
    End With
    
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================

    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub


'���e:
'   ���[�U�}�N�������������ŃR�[������i��������΁j
'
'�p�����[�^:
'    [Arg0]      In   ���[�U�[�}�N����
'
'�߂�l:
'
'���ӎ���:

Public Sub FW_CallUserMacro(ByVal Parameter As CSetFunctionInfo)
    
On Error GoTo ErrHandler

    If Parameter.ArgParameterCount <> 1 Then
         Err.Raise 9999, "FW_CallUserMacro", "The number of FW_CallUserMacro's arguments is invalid." & " @ " & Parameter.ConditionName
    End If

    Call Application.Run(Parameter.Arg(0))
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'��
'��������������������������������������
'subCurrent_Serial_Test�p:Start
'��������������������������������������
'��

'���e:
'   SUB�d������̃p�����[�^�ݒ���s��
'
'�p�����[�^:
'    [Arg0]      In   ���ω�
'    [Arg1]      In   �N�����v�d��(A)
'    [Arg2]      In   WaitTime(s)
'    [Arg3]      In   �s�����\�[�X
'
'�߂�l:
'
'���ӎ���:
'2012/10/19 DC_WG �C��
'2013/02/01 MB_WG �C��
'2013/09/27 �ύX
'
Public Sub FW_SetSubCurrentParam(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 4 Then
         Err.Raise 9999, "FW_SetSubCurrentParam", "The number of FW_SetSubCurrentParam's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '========Check Average Count============================================
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_SetSubCurrentParam", "FW_SetSubCurrentParam Arg0: Type Mismatch """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    Dim lCount As Long
    lCount = Parameter.Arg(0)
    
    '========Check Clamp Current============================================
    If Not IsNumeric(Parameter.Arg(1)) Then
        Err.Raise 9999, "FW_SetSubCurrentParam", "FW_SetSubCurrentParam Arg1: Type Mismatch """ & Parameter.Arg(1) & """ @ " & Parameter.ConditionName
    End If
    
    Dim dblClampCurrent As Double
    dblClampCurrent = Parameter.Arg(1)
    
    '========Check Wait Time ===============================================
    If Not IsNumeric(Parameter.Arg(2)) Then
        Err.Raise 9999, "FW_SetSubCurrentParam", "FW_SetSubCurrentParam Arg2: Type Mismatch """ & Parameter.Arg(2) & """ @ " & Parameter.ConditionName
    End If
    
    Dim dblWaitTime As Double
    dblWaitTime = Parameter.Arg(2)
    
    Dim strResourceName As String
    If UCase(Parameter.Arg(3)) = "BPMU" Then
        strResourceName = "BPMU"
    Else
        strResourceName = "Not BPMU"
    End If
    
    '========Add SubCurrentParam To VarBank====================================
    Dim strCountKey As String, strClampKey As String, strWaitTimeKey As String, strPinResourceKey As String
    strCountKey = GetInstanceName & SUBCURRENT_AVERAGE_COUNT
    strClampKey = GetInstanceName & SUBCURRENT_CLAMP_CURRENT
    strWaitTimeKey = GetInstanceName & SUBCURRENT_WAIT_TIME
    strPinResourceKey = GetInstanceName & SUBCURRENT_PIN_RESOURCE

    With TheVarBank
        If .IsExist(strCountKey) = True Then
            Err.Raise 9999, "FW_SetSubCurrentParam", "FW_SetSubCurrentParam was already called for AverageCount: "
        ElseIf .IsExist(strClampKey) = True Then
            Err.Raise 9999, "FW_SetSubCurrentParam", "FW_SetSubCurrentParam was already called for ClampCurrent: "
        ElseIf .IsExist(strWaitTimeKey) = True Then
            Err.Raise 9999, "FW_SetSubCurrentParam", "FW_SetSubCurrentParam was already called for WaitTime: "
        ElseIf .IsExist(strPinResourceKey) = True Then
            Err.Raise 9999, "FW_SetSubCurrentParam", "FW_SetSubCurrentParam was already called for PinResource: "
        Else
            Call .Add(strCountKey, lCount, False, strCountKey)
            Call .Add(strClampKey, dblClampCurrent, False, strClampKey)
            Call .Add(strWaitTimeKey, dblWaitTime, False, strWaitTimeKey)
            Call .Add(strPinResourceKey, strResourceName, False, strPinResourceKey)
        End If
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub
'���e:
'   SUB�d������̃p�����[�^�擾
'
'�p�����[�^:
'
'�߂�l:
'       ���ω�
'
'���ӎ���:
'
Public Function GetSubCurrentAverageCount(ByVal strInstanceName As String) As Long

    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & SUBCURRENT_AVERAGE_COUNT) Then
        Err.Raise 9999, "GetSubCurrentAverageCount", "FW_SetSubCurrentParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetSubCurrentAverageCount = TheVarBank.Value(strInstanceName & SUBCURRENT_AVERAGE_COUNT)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Function

'���e:
'   SUB�d������̃p�����[�^�擾
'
'�p�����[�^:
'
'�߂�l:
'       �N�����v�d���l(A)
'
'���ӎ���:
'
Public Function GetSubCurrentClampCurrent(ByVal strInstanceName As String) As Double

    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & SUBCURRENT_CLAMP_CURRENT) Then
        Err.Raise 9999, "GetSubCurrentClampCurrent", "FW_SetSubCurrentParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetSubCurrentClampCurrent = TheVarBank.Value(strInstanceName & SUBCURRENT_CLAMP_CURRENT)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�

End Function

'���e:
'   SUB�d������̃p�����[�^�擾
'
'�p�����[�^:
'
'�߂�l:
'       WaitTime(s)
'
'���ӎ���:
'
Public Function GetSubCurrentWaitTime(ByVal strInstanceName As String) As Double

    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & SUBCURRENT_WAIT_TIME) Then
        Err.Raise 9999, "GetSubCurrentWaitTime", "FW_SetSubCurrentParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetSubCurrentWaitTime = TheVarBank.Value(strInstanceName & SUBCURRENT_WAIT_TIME)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�

End Function
'���e:
'   �L���v�`���p�����[�^�̐ݒ���s��
'
'�p�����[�^:
'    [Arg0]      In   ���ω�
'    [Arg1]      In   ���ω����[�h(Average or NoAverage)
'
'�߂�l:
'
'���ӎ���:
'
'2013/01/22 H.Arikawa Arg3 -> Arg21�֕ύX
Public Sub FW_SetCaptureParam(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler:
    
    '========Param check  ==================================================
    If Parameter.ArgParameterCount() <> 4 Then
        Err.Raise 9999, "FW_SetCaptureParam", "The Number of arguments is invalid! """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    
    '========Check Average Count============================================
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_SetCaptureParam", "Arg0: Type Mismatch """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    Dim lCount As Long
    lCount = Parameter.Arg(0)
    If (lCount < 1) Or (512 < lCount) Then 'Check For CaptureUnit
        Err.Raise 9999, "FW_SetCaptureParam", "Arg0: Range Invalid """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    '========Check Average Mode============================================
    Dim strAverageMode As String
    strAverageMode = Parameter.Arg(1)
    If lCount = 1 And strAverageMode = "NonAverage" Then strAverageMode = "Average"
    
    If strAverageMode <> CAPTURE_AVERAGE_MODE_AVERAGE And strAverageMode <> CAPTURE_AVERAGE_MODE_NO_AVERAGE Then
        Err.Raise 9999, "FW_SetCaptureParam", " Arg1: Value Invalid  """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    '========Check Frame Skip Count============================================
    If Not IsNumeric(Parameter.Arg(2)) Then
        Err.Raise 9999, "FW_SetCaptureParam", "Arg2: Type Mismatch """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    Dim lSkipCount As Long
    lSkipCount = Parameter.Arg(2)
    If (lCount < 0) Or (512 < lCount) Then 'Check For CaptureUnit
        Err.Raise 9999, "FW_SetCaptureParam", "Arg2: Range Invalid [0-512] """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    Dim acqInstances() As String
    acqInstances = Split(Parameter.Arg(3), ",")
    Dim i As Long
    '========Add Average Set To VarBank====================================
    Dim strCountKey As String, strModekey As String, strSkipKey As String
    
    For i = 0 To UBound(acqInstances)
        strCountKey = acqInstances(i) & CAPTURE_PARAM_AVERAGE_COUNT
        strModekey = acqInstances(i) & CAPTURE_PARAM_AVERAGE_MODE
        strSkipKey = acqInstances(i) & CAPTURE_PARAM_FRAME_SKIP
    
        With TheVarBank
            If .IsExist(strCountKey) = True Then
                If .Value(strCountKey) <> lCount Then
                    Err.Raise 9999, "FW_SetCaptureParam", "FW_SetCaptureParam already called for AverageCount  """ & .Value(strCountKey) & """ @ " & Parameter.ConditionName
                Else
                    '�Ȃɂ����Ȃ�
                End If
            Else
                Call .Add(strCountKey, lCount, False, strCountKey)
            End If
            
            If .IsExist(strModekey) = True Then
                If .Value(strModekey) <> strAverageMode Then
                    Err.Raise 9999, "FW_SetCaptureParam", "FW_SetCaptureParam already called for AverageMode  """ & .Value(strCountKey) & """ @ " & Parameter.ConditionName
                    '�Ȃɂ����Ȃ�
                End If
            Else
                Call .Add(strModekey, strAverageMode, False, strModekey)
            End If
            
            If .IsExist(strSkipKey) = True Then
                If .Value(strSkipKey) <> lSkipCount Then
                    Err.Raise 9999, "FW_SetCaptureParam", "FW_SetCaptureParam already called for FrameSkip  """ & .Value(strCountKey) & """ @ " & Parameter.ConditionName
                    '�Ȃɂ����Ȃ�
                End If
            Else
                Call .Add(strSkipKey, lSkipCount, False, strSkipKey)
            End If
            
        End With
    Next i
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �摜�L���v�`���̃p�����[�^�擾
'
'�p�����[�^:
'
'�߂�l:
'       ���ω�
'
'���ӎ���:
'
Public Function GetCaptureParamAverageCount(ByVal strInstanceName As String) As Long
    
    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & CAPTURE_PARAM_AVERAGE_COUNT) Then
        Err.Raise 9999, "GetCaptureParamAverageCount", "FW_SetCaptureParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetCaptureParamAverageCount = TheVarBank.Value(strInstanceName & CAPTURE_PARAM_AVERAGE_COUNT)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Function

'���e:
'   �摜�L���v�`���̃p�����[�^�擾
'
'�p�����[�^:
'
'�߂�l:
'       ���ω����[�h
'
'���ӎ���:
'
Public Function GetCaptureParamAverageMode(ByVal strInstanceName As String) As IdpAverageMode
    
    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & CAPTURE_PARAM_AVERAGE_MODE) Then
        Err.Raise 9999, "GetCaptureParamAverageMode", "FW_SetCaptureParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    With TheVarBank
        If (.Value(strInstanceName & CAPTURE_PARAM_AVERAGE_MODE) = CAPTURE_AVERAGE_MODE_AVERAGE) Then
            GetCaptureParamAverageMode = idpAverage
            Exit Function
        End If
         If (.Value(strInstanceName & CAPTURE_PARAM_AVERAGE_MODE) = CAPTURE_AVERAGE_MODE_NO_AVERAGE) Then
            GetCaptureParamAverageMode = idpNonAverage
            Exit Function
        End If
   End With
   
   Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Function

'���e:
'   �摜�L���v�`���̃p�����[�^�擾
'
'�p�����[�^:
'
'�߂�l:
'       FrameSkip��
'
'���ӎ���:
'
Public Function GetCaptureParamFrameSkip(ByVal strInstanceName As String) As Long
    
    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & CAPTURE_PARAM_FRAME_SKIP) Then
        Err.Raise 9999, "GetCaptureParamFrameSkip", "FW_SetCaptureParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetCaptureParamFrameSkip = TheVarBank.Value(strInstanceName & CAPTURE_PARAM_FRAME_SKIP)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Function

'���e:
'   SCRN����SetMV����MeasureV�܂ł�Wait���Ԑݒ�
'
'�p�����[�^:
'    [Arg0]      In   Average Count
'    [Arg1]      In   WaitTime(s)
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_SetScrnMeasureParam(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 2 Then
         Err.Raise 9999, "FW_SetScrnMeasureParam", "The number of FW_SetScrnMeasureParam's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '========Average ============================================
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_SetScrnMeasureParam", "FW_SetScrnMeasureParam Arg0: Type Mismatch """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    Dim lAverage As Long
    lAverage = Parameter.Arg(0)
    
    
    '========Wait Time ============================================
    If Not IsNumeric(Parameter.Arg(1)) Then
        Err.Raise 9999, "FW_SetScrnMeasureParam", "FW_SetScrnMeasureParam Arg0: Type Mismatch """ & Parameter.Arg(1) & """ @ " & Parameter.ConditionName
    End If
    
    Dim dblWaitTime As Double
    dblWaitTime = Parameter.Arg(1)
    
    
    '========Add SubCurrentParam To VarBank====================================
    Dim strCountKey As String
    strCountKey = GetInstanceName & SCRN_MEAS_AVERAGE_COUNT
    Dim strWaitTimeKey As String
    strWaitTimeKey = GetInstanceName & SCRN_MEAS_WAIT_TIME

    With TheVarBank
        If .IsExist(strCountKey) = True Then
            Err.Raise 9999, "FW_ScrnMeasureWaitParam", "FW_ScrnMeasureWaitParam was already called for WaitTime: "
        Else
            Call .Add(strCountKey, lAverage, False, strCountKey)
        End If
    End With
    
    With TheVarBank
        If .IsExist(strWaitTimeKey) = True Then
            Err.Raise 9999, "FW_ScrnMeasureWaitParam", "FW_ScrnMeasureWaitParam was already called for WaitTime: "
        Else
            Call .Add(strWaitTimeKey, dblWaitTime, False, strWaitTimeKey)
        End If
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   SCRN�̃p�����[�^�擾
'
'�p�����[�^:
'
'�߂�l:
'       SetMV����Measure�܂ł�Wait���Ԏ擾
'
'���ӎ���:
'
Public Function GetScrnMeasureWaitTime(ByVal strInstanceName As String) As Double
    
    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & SCRN_MEAS_WAIT_TIME) Then
        Err.Raise 9999, "GetScrnMeasureWaitParam", "GetScrnMeasureWaitParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetScrnMeasureWaitTime = TheVarBank.Value(strInstanceName & SCRN_MEAS_WAIT_TIME)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Function


'���e:
'   SCRN�̃p�����[�^�擾
'
'�p�����[�^:
'
'�߂�l:
'       SCRN������̕��ω�
'
'���ӎ���:
'
Public Function GetScrnMeasureAverageCount(ByVal strInstanceName As String) As Double
    
    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & SCRN_MEAS_AVERAGE_COUNT) Then
        Err.Raise 9999, "GetScrnMeasureAverageCount", "GetScrnMeasureAverageCount in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetScrnMeasureAverageCount = TheVarBank.Value(strInstanceName & SCRN_MEAS_AVERAGE_COUNT)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Function
'���e:
'   To execute wait for screening.
'
'�p�����[�^:
'   [Arg0]  In  Screening wait time in second, specified on the specification sheet.
'   [Arg1]  In  Wait time between "SET" and "MEASUREMENT" for the dc test
'   [Arg1]  In  WaitTime(s)
'
'�߂�l:
'
'���ӎ���:
'
Public Sub FW_ScreeningWait(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    Const NUMBER_OF_ARGUMENTS As Long = 3
    If Parameter.ArgParameterCount <> NUMBER_OF_ARGUMENTS Then
         Err.Raise 9999, "FW_ScreeningWait", "The number of FW_ScreeningWait's arguments must be " & NUMBER_OF_ARGUMENTS & "." & " @ " & Parameter.ConditionName
         GoTo ErrHandler
    End If
    
    '======== specification wait time ============================================
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_ScreeningWait", "FW_ScreeningWait Arg0: Type Mismatch (must be numeric) """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
        GoTo ErrHandler
    End If
    
    Dim dblScreeningWait As Double
    dblScreeningWait = Parameter.Arg(0)
    
    
    '========DC measurement wait Time ============================================
    If Not IsNumeric(Parameter.Arg(1)) Then
        If Parameter.Arg(1) = "-" Then
            Parameter.Arg(1) = 0
        Else
            Err.Raise 9999, "FW_ScreeningWait", "FW_ScreeningWait Arg0: Type Mismatch (must be numeric or " - " ) """ & Parameter.Arg(1) & """ @ " & Parameter.ConditionName
            GoTo ErrHandler
        End If
    End If
    
    Dim dblDcWaitTime As Double
    dblDcWaitTime = Parameter.Arg(1)
    
    '========TOPT mode ============================================
    Dim isToptMode As Boolean
    isToptMode = Parameter.Arg(2)
    
    'Wait����
    Dim dblTotalWaitTime As Double
    If dblScreeningWait > dblDcWaitTime Then dblTotalWaitTime = dblScreeningWait - dblDcWaitTime
    If TheExec.RunOptions.AutoAcquire = True And isToptMode Then
        Call TheHdw.TOPT.WAIT(toptTimer, dblTotalWaitTime * 1000)
    Else
        Call TheHdw.WAIT(dblTotalWaitTime)
    End If
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub


'���e:
'   �s����GND�ɐڑ�����(�ڒn)
'
'�p�����[�^:
'    [Arg0]      In   �s�����i�s���O���[�v���j
'
'�߂�l:
'
'���ӎ���:
'       �SSite��1�s������GND�ɐڑ��B
'
Public Sub FW_SetGND(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 1 Then
        Err.Raise 9999, "FW_SetGND", "The number of FW_SetGND's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_TESTER_CHANNELS
        Call CheckBeforeTestCondition(eMode, Parameter)
    End If
'=========Before TestCondition Check End ========================
    Call SetGND(Parameter.Arg(0))
'=========After TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Call CheckAfterTestCondition(eMode, Parameter)
    End If
'=========After TestCondition Check End ========================
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   SUB�d������̃p�����[�^�擾
'
'�p�����[�^:
'
'�߂�l:
'       PinResourceName
'
'���ӎ���:
'     2012/10/19 DC_WG �ǉ�
'     2012/11/1  Stop Delete
'
Public Function GetSubCurrentPinResourceName(ByVal strInstanceName As String) As String

    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & SUBCURRENT_PIN_RESOURCE) Then
        Err.Raise 9999, "GetSubCurrentPinResourceName", "FW_SetSubCurrentParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetSubCurrentPinResourceName = TheVarBank.Value(strInstanceName & SUBCURRENT_PIN_RESOURCE)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�

End Function
'��
'��������������������������������������
'subCurrent_Serial_Test�p:End
'��������������������������������������
'��

'��
'��������������������������������������
'C�B_OV���ڗp�����[�ݒ�:Start
'��������������������������������������
'��

'���e:
'   �����[ON���s��
'
'�p�����[�^:
'    [Arg0]      In   �����[UB
'
'�߂�l:
'
'���ӎ���:S�B�ł͎g�p���Ȃ����A���Q�͂Ȃ��ׁA���ʂ�ConditionMacro�Ƃ���ׂɓ���Ă����B
'     2012/11/1  Stop Delete
'
'''Public Sub FW_SET_RELAY_ON(ByVal Parameter As CSetFunctionInfo)
'''
'''    On Error GoTo ErrHandler
'''
'''    If Parameter.ArgParameterCount() <> 1 Then
'''        err.Raise 9999, "FW_SET_RELAY_ON", "The number of FW_SET_RELAY_ON's arguments is invalid." & " @ " & Parameter.ConditionName
'''    End If
'''
''''=========Before TestCondition Check Start ======================
'''#If EEEAUTO_AUTO_MODIFY_TESTCONDITION = 1 Then
'''    Dim eMode As eTestCnditionCheck
'''    eMode = APMU_RELAY_UB_ON
'''    Call CheckBeforeTestCondition(eMode, Parameter)
'''#End If
''''=========Before TestCondition Check End ========================
'''
'''    'RELAY_ON
'''    DutConnectDbNumber Parameter.Arg(0), True
'''
''''=========After TestCondition Check Start ======================
'''#If EEEAUTO_AUTO_MODIFY_TESTCONDITION = 1 Then
'''    Call CheckAfterTestCondition(eMode, Parameter)
'''#End If
''''=========After TestCondition Check End ========================
'''
'''    Exit Sub
'''
'''ErrHandler:
'''    MsgBox "Error Occured !! " & CStr(err.Number) & " - " & err.Source & chR(13) & chR(13) & err.Description
'''    Call DisableAllTest 'EeeJob�֐�
'''
'''End Sub

'���e:
'   �����[OFF���s��
'
'�p�����[�^:
'    [Arg0]      In   �����[UB
'
'�߂�l:
'
''''���ӎ���:S�B�ł͎g�p���Ȃ����A���Q�͂Ȃ��ׁA���ʂ�ConditionMacro�Ƃ���ׂɓ���Ă����B
''''     2012/11/1  Stop Delete
''''
'''Public Sub FW_SET_RELAY_OFF(ByVal Parameter As CSetFunctionInfo)
'''
'''    On Error GoTo ErrHandler
'''
'''    If Parameter.ArgParameterCount() <> 1 Then
'''        err.Raise 9999, "FW_SET_RELAY_OFF", "The number of FW_SET_RELAY_OFF's arguments is invalid." & " @ " & Parameter.ConditionName
'''    End If
'''
''''=========Before TestCondition Check Start ======================
'''#If EEEAUTO_AUTO_MODIFY_TESTCONDITION = 1 Then
'''    Dim eMode As eTestCnditionCheck
'''    eMode = APMU_RELAY_UB_OFF
'''    Call CheckBeforeTestCondition(eMode, Parameter)
'''#End If
''''=========Before TestCondition Check End ========================
'''
'''    'RELAY_OFF
'''    DutConnectDbNumber Parameter.Arg(0), False
'''
''''=========After TestCondition Check Start ======================
'''#If EEEAUTO_AUTO_MODIFY_TESTCONDITION = 1 Then
'''    Call CheckAfterTestCondition(eMode, Parameter)
'''#End If
''''=========After TestCondition Check End ========================
'''
'''    Exit Sub
'''
'''ErrHandler:
'''    MsgBox "Error Occured !! " & CStr(err.Number) & " - " & err.Source & chR(13) & chR(13) & err.Description
'''    Call DisableAllTest 'EeeJob�֐�
'''
'''End Sub
'''
'''Public Function DutConnectDbNumber(ByVal dbNum As Long, ByVal ONOFF As Boolean)
'''
'''    TheHdw.APMU.board(APMU_BOARD_NUMBER).UtilityBit(dbNum) = ONOFF '2012/11/15 175Debug Arikawa
'''
'''End Function
'''''''��
'��������������������������������������
'C�B_OV���ڗp�����[�ݒ�:End
'��������������������������������������
'��
'��
'��������������������������������������
'Hold_Voltage_Test�p:Start
'��������������������������������������
'��

'���e:
'   HOLD�d������̃p�����[�^�ݒ���s��
'
'�p�����[�^:
'    [Arg0]      In   ���ω�
'    [Arg1]      In   �N�����v�d��(A)
'    [Arg2]      In   WaitTime(s)
'
'�߂�l:
'
'���ӎ���:
'     2012/11/1  Stop Delete

Public Sub FW_SetHoldVoltageParam(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount <> 3 Then
         Err.Raise 9999, "FW_SetHoldVoltageParam", "The number of FW_SetHoldVoltageParam's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '========Check Average Count============================================
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_SetHoldVoltageParam", "FW_SetHoldVoltageParam Arg0: Type Mismatch """ & Parameter.Arg(0) & """ @ " & Parameter.ConditionName
    End If
    
    Dim lCount As Long
    lCount = Parameter.Arg(0)
    
    '========Check Clamp Current============================================
    If Not IsNumeric(Parameter.Arg(1)) Then
        Err.Raise 9999, "FW_SetHoldVoltageParam", "FW_SetHoldVoltageParam Arg1: Type Mismatch """ & Parameter.Arg(1) & """ @ " & Parameter.ConditionName
    End If
    
    Dim dblClampCurrent As Double
    dblClampCurrent = Parameter.Arg(1)
    
    '========Check Wait Time ===============================================
    If Not IsNumeric(Parameter.Arg(2)) Then
        Err.Raise 9999, "FW_SetHoldVoltageParam", "FW_SetHoldVoltageParam Arg2: Type Mismatch """ & Parameter.Arg(2) & """ @ " & Parameter.ConditionName
    End If
    
    Dim dblWaitTime As Double
    dblWaitTime = Parameter.Arg(2)
    
    '========Add HoldVoltageParam To VarBank====================================
    Dim strCountKey As String, strClampKey As String, strWaitTimeKey As String
    strCountKey = GetInstanceName & HOLDVOLTAGE_AVERAGE_COUNT
    strClampKey = GetInstanceName & HOLDVOLTAGE_CLAMP_CURRENT
    strWaitTimeKey = GetInstanceName & HOLDVOLTAGE_WAIT_TIME

    With TheVarBank
        If .IsExist(strCountKey) = True Then
            Err.Raise 9999, "FW_SetHoldVoltageParam", "FW_SetHoldVoltageParam was already called for AverageCount: "
        ElseIf .IsExist(strClampKey) = True Then
            Err.Raise 9999, "FW_SetHoldVoltageParam", "FW_SetHoldVoltageParam was already called for ClampCurrent: "
        ElseIf .IsExist(strWaitTimeKey) = True Then
            Err.Raise 9999, "FW_SetHoldVoltageParam", "FW_SetHoldVoltageParam was already called for WaitTime: "
        Else
            Call .Add(strCountKey, lCount, False, strCountKey)
            Call .Add(strClampKey, dblClampCurrent, False, strClampKey)
            Call .Add(strWaitTimeKey, dblWaitTime, False, strWaitTimeKey)
        End If
    End With
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   HOLD�d������̃p�����[�^�擾
'
'�p�����[�^:
'
'�߂�l:
'       ���ω�
'
'���ӎ���:
'     2012/11/1  Stop Delete

Public Function GetHoldVoltageAverageCount(ByVal strInstanceName As String) As Long

    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & HOLDVOLTAGE_AVERAGE_COUNT) Then
        Err.Raise 9999, "GetHoldVoltageAverageCount", "FW_SetHoldVoltageParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetHoldVoltageAverageCount = TheVarBank.Value(strInstanceName & HOLDVOLTAGE_AVERAGE_COUNT)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Function

'���e:
'   HOLD�d������̃p�����[�^�擾
'
'�p�����[�^:
'
'�߂�l:
'       �N�����v�d���l(A)
'
'���ӎ���:
'     2012/11/1  Stop Delete

Public Function GetHoldVoltageClampCurrent(ByVal strInstanceName As String) As Double

    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & HOLDVOLTAGE_CLAMP_CURRENT) Then
        Err.Raise 9999, "GetHoldVoltageClampCurrent", "FW_SetHoldVoltageParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetHoldVoltageClampCurrent = TheVarBank.Value(strInstanceName & HOLDVOLTAGE_CLAMP_CURRENT)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�

End Function

'���e:
'   HOLD�d������̃p�����[�^�擾
'
'�p�����[�^:
'
'�߂�l:
'       WaitTime(s)
'
'���ӎ���:
'     2012/11/1  Stop Delete

Public Function GetHoldVoltageWaitTime(ByVal strInstanceName As String) As Double

    On Error GoTo ErrHandler:
    
    '========Error Check====================================
    If Not TheVarBank.IsExist(strInstanceName & HOLDVOLTAGE_WAIT_TIME) Then
        Err.Raise 9999, "GetHoldVoltageWaitTime", "FW_SetHoldVoltageParam in not called for " & strInstanceName
    End If
    
    '=======Get Value From VarBank =========================
    GetHoldVoltageWaitTime = TheVarBank.Value(strInstanceName & HOLDVOLTAGE_WAIT_TIME)
    Exit Function
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�

End Function
'��
'��������������������������������������
'Hold_Voltage_Test�p:End
'��������������������������������������
'��

'��
'��������������������������������������
'DC TOPT�p�@FW_DcTopt_Set:
'��������������������������������������
'��
'

'���e:
'   DC TOPT���p���ɁASET���̃V�i���I���Ăяo���ADC����̈��艻���ԕ�TOPT Wait���Ăяo��
'
'�p�����[�^:
'    [Arg0]      In DC Test Scenario Name�BSET���̃V�i���I��
'    [Arg1]     DC Test Scenario��SET���V�i���I��MEASURE���V�i���I�̊Ԃ�
'               Wait���ԁB
'
'�߂�l:
'
'���ӎ���:
'     DC TOPT�̏ꍇ�ASET��MEASURE�Ԃ�Wait���Ԃ�DC Test Scenario�ł͐��䂹���A
'   �����̒l�Ő��䂵�܂��B�]���āA�f�o�b�O���ɂ́ATest Condition�̖{�֐�
'   ��Arg1�̒l��ҏW���Ă��������B
'
Public Sub FW_DcTopt_Set(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 2 Then
        Err.Raise 9999, "FW_DcTopt_Set", "The number of FW_DcTopt_Set's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
                        
    '========== DC�V�i���I�V�[�g���s ===============================
    TheDcTest.SetScenario (Parameter.Arg(0))
    TheDcTest.Execute
        
    '========= Wait ======================
    If TheExec.RunOptions.AutoAcquire = True Then
        Call TheHdw.TOPT.WAIT(toptTimer, Parameter.Arg(1) * 1000)
    Else
        Call TheHdw.WAIT(Parameter.Arg(1))
    End If
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'��
'��������������������������������������
'DC TOPT�p�@FW_DcTopt_Measure:
'��������������������������������������
'��
'

'���e:
'   DC TOPT���p���ɁAMEASURE���̃V�i���I���Ăяo���B
'
'�p�����[�^:
'    [Arg0]      In DC Test Scenario Name�BMEASURE���̃V�i���I��
'
'�߂�l:
'
'���ӎ���:
'   DC TOPT���p���ɁADC Test Scenario�V�[�g�ŁAMEASURE���̃V�i���I��
'   Wait���L�ڂ��Ă��L���ɂ͂Ȃ�܂���B����́ADC Test Scenario�V�[�g�ł́A
'   "MEASURE"�A�N�V������Wait�́A����V�i���I�̒��O�Ɏ��s����Ă���"SET"
'   �A�N�V���������Wait���ԂƂ��Ĉ����邽�߂ł��B
'   DC TOPT���p���ɂ́A�V�i���I��"MEASURE"�A�N�V��������J�n���A���O��"SET"
'   �A�N�V����������܂���̂ŁAWait���Ԃ��L�ڂ��Ă���������܂��B
'
Public Sub FW_DcTopt_Measure(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 1 Then
        Err.Raise 9999, "FW_DcTopt_Measure", "The number of FW_DcTopt_Measure's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
                        
    '========== DC�V�i���I�V�[�g���s ===============================
    TheDcTest.SetScenario (Parameter.Arg(0))
    TheDcTest.Execute
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'��
'��������������������������������������
'SCRN TOPT�p�@FW_ScrnWaitSet:
'��������������������������������������
'��
'

'���e:
'   �w�莞�ԕ�Wait������
'
'�p�����[�^:
'    [Arg0]      In   Wait����(s)
'
'�߂�l:
'
'���ӎ���:
'     TheHdw.Wait�Ŗⓚ���p�ɑ҂�
'     2012/11/1  Stop Delete
'
Public Sub FW_ScrnWaitSet(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 1 Then
        Err.Raise 9999, "FW_ScrnWaitSet", "The number of FW_ScrnWaitSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_ScrnWaitSet", "FW_ScrnWaitSet's Arg0 is invalid type." & " @ " & Parameter.ConditionName
    End If
    
    Dim dblWaitTime As Double
    dblWaitTime = Parameter.Arg(0)
    
    'Wait����
    Call TheHdw.WAIT(dblWaitTime)
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'��
'��������������������������������������
'SCRN TOPT�p�@FW_ScrnWaitSetTopt:
'��������������������������������������
'��
'

'���e:
'   �w�莞�ԕ�Wait������
'
'�p�����[�^:
'    [Arg0]      In   Wait����(s)
'
'�߂�l:
'
'���ӎ���:
'     TheExec.RunOptions.AutoAcquire
'     �ɉ�����Wait�𕪂���A�Ăяo������TOPT���s���Ă��邩�ӎ�����K�v������
'     TOPT���s���łȂ���Ί��҂�������͂��Ȃ�
'     2012/11/1  Stop Delete
'
Public Sub FW_ScrnWaitSetTopt(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 1 Then
        Err.Raise 9999, "FW_ScrnWaitSetTopt", "The number of FW_ScrnWaitSetTopt's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_ScrnWaitSetTopt", "FW_ScrnWaitSetTopt's Arg0 is invalid type." & " @ " & Parameter.ConditionName
    End If
    
    Dim dblWaitTime As Double
    dblWaitTime = Parameter.Arg(0)
    
    'Wait����
    If TheExec.RunOptions.AutoAcquire = True Then
        Call TheHdw.TOPT.WAIT(toptTimer, dblWaitTime * 1000)
    Else
        Call TheHdw.WAIT(dblWaitTime)
    End If
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub
Public Sub FW_PowerDownAndDisconnectPins(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler

    If Parameter.ArgParameterCount() <> 0 Then
        Err.Raise 9999, "FW_PowerDownAndDisconnectPins", "The number of arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Call PowerDownAndDisconnect
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

Public Sub FW_PowerDown(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler

    If Parameter.ArgParameterCount() <> 0 Then
        Err.Raise 9999, "FW_PowerDown", "The number of arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Call PowerDown4ApmuUnderShoot
    
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �p�^�[���������Ă��邩�ǂ����X�e�[�^�X���擾����B
'
'�p�����[�^:
'
'
'�߂�l:
'
'���ӎ���:Halt���g�p���Ă���p�^�[���ł̂ݎg�p����B(������Loop�ɂȂ��)
'         keep_alive�g�p�^�C�v�͗v�m�F(12/20)
'GUI�Ń��[�U�[��TOPT�L���JOB������I�������ꍇ��FW_PatSetTypeSelect�ƃZ�b�g�ɐ��������Condition�B
'

Public Sub FW_PatStatus(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
    If Parameter.ArgParameterCount() <> 0 Then
        Err.Raise 9999, "FW_PatStatus", "The number of FW_PatStatus's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    If TheExec.RunOptions.AutoAcquire = True Then
        Dim iStatus As Long
        If TheHdw.Digital.Patgen.IsRunningAnySite = True Then      'True:Still Running
            iStatus = 0
        ElseIf TheHdw.Digital.Patgen.IsRunningAnySite = False Then 'False:haltend or keepalive
            iStatus = 1
        End If
        
        While (iStatus <> 1)
            If PatCheckCounter < 999 Then
                TheHdw.TOPT.Recall
                PatCheckCounter = PatCheckCounter + 1
                Call WaitSet(10 * mS)
                Exit Sub
            End If
        Wend
    End If

    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �p�^���̊J�n�������Ȃ�(�I����҂��Ȃ�/�I����҂�)
'
'�p�����[�^:
'    [Arg0]      In   �p�^����
'    [Arg1]      In   TSB�V�[�g��
'    [Arg2]      In   ���s��̃E�F�C�g�^�C��(�ȗ��\ �ȗ���Wait�Ȃ�)
'
'�߂�l:
'
'GUI�Ń��[�U�[��TOPT�L���JOB������I�������ꍇ��PatRun�̑���ɐ��������Condition�B
'
Public Sub FW_PatSetTypeSelect(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
    Const PAT_START_LABEL As String = "START"

    If Parameter.ArgParameterCount <> 2 And Parameter.ArgParameterCount <> 3 Then
         Err.Raise 9999, "FW_PatSet", "The number of FW_PatSet's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPatGroupName As String
    Dim strTsbName As String
        
    With Parameter
        strPatGroupName = .Arg(0)
        strTsbName = .Arg(1)
    End With
    
    Call StopPattern_Halt 'EeeJob�֐�
    Call SetTimeOut 'EeeJob�֐�
    
    'TOPT�̎g�p�ɉ�����PatRun/PatSet�����I������B
    If TheExec.RunOptions.AutoAcquire = True Then
        With TheHdw.Digital
            Call .Timing.Load(strTsbName)
            Call .Patterns.pat(strPatGroupName).Start(PAT_START_LABEL)
        End With
    ElseIf TheExec.RunOptions.AutoAcquire = False Then
        With TheHdw.Digital
            Call .Timing.Load(strTsbName)
            Call .Patterns.pat(strPatGroupName).Run(PAT_START_LABEL)
        End With
    End If
    
   '�҂����Ԃ̎w�肪����ꍇ�A�҂�
    If Parameter.ArgParameterCount() = 3 Then
        Dim dblWaitTime As Double
        If Not IsNumeric(Parameter.Arg(2)) Then
            Err.Raise 9999, "FW_PatSet", "FW_PatSet's Arg2 is invalid type." & " @ " & Parameter.ConditionName
        End If
        dblWaitTime = Parameter.Arg(2)
        Call TheHdw.WAIT(dblWaitTime)
    End If
    
    PatCheckCounter = 0
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'01/29 Add H.Arikawa
'�����̏����ݒ�ȗ��̏������ɓ�������肷��B(������ނɉ�����)
'�b�菈��������B

Public Sub FW_OptEscape(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 2 Then
        Err.Raise 9999, "FW_OptEscape", "The number of FW_OptEscape's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------
    
    'Call "OptIni" in case the following object "OptCond" is nothing. (Optcond�I�u�W�F�N�g��Nothing��������OptIni��������)
    Call sub_CheckOptCond
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_ILLUMINATOR_ESCAPSE
        
        Call CheckBeforeTestCondition(eMode, Parameter)
        If Not IsValidTestCondition(eMode, Parameter) Then Exit Sub
        
    End If
'=========Before TestCondition Check End ========================
    
    OptCheckCounter = 0
        
    '�����ݒ� Escape
    'NSIS3/3KAI : PIN�ɑޔ�����B
    'NSIS5/5KAI : Up�ɑޔ�����B
    
    If OptCond.IllumMaker = NIKON Then
        With Parameter
            If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS3KAI" Then
                Call OptMod("PIN", .Arg(0))
            ElseIf OptCond.IllumModel = "N-SIS5" Or OptCond.IllumModel = "N-SIS5KAI" Then
                Call OptModZ_NSIS5("Up", .Arg(0))
            End If
        End With
    End If
    
    Exit Sub
   
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'01/29 Add H.Arikawa
'�����̏����ݒ�ȗ��̏������ɓ�������肷��B(������ނɉ�����)
'�b�菈��������B

Public Sub FW_OptModOrModZ1(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 2 Then
        Err.Raise 9999, "FW_OptEscape", "The number of FW_OptEscape's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------
    
    'Call "OptIni" in case the following object "OptCond" is nothing. (Optcond�I�u�W�F�N�g��Nothing��������OptIni��������)
    Call sub_CheckOptCond
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_ILLUMINATOR_MODZ1
        Call CheckBeforeTestCondition(eMode, Parameter)
        If Not IsValidTestCondition(eMode, Parameter) Then Exit Sub
    
    End If
'=========Before TestCondition Check End ========================

    OptCheckCounter = 0
        
    'T
    'PIN   NSIS-5 Escape Point
    'F_UNIT
    
    '�����ݒ� ModOrModZ1
    If OptCond.IllumMaker = NIKON Then
        With Parameter
            If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS5" Or OptCond.IllumModel = "N-SIS5KAI" Then
                Call OptMod(.Arg(1), .Arg(0))
            ElseIf OptCond.IllumModel = "N-SIS3KAI" Then
                Call OptModZ_NSIS5(.Arg(1), .Arg(0))
            End If
        End With
    End If
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'01/29 Add H.Arikawa
'�����̏����ݒ�ȗ��̏������ɓ�������肷��B(������ނɉ�����)
'�b�菈��������B

Public Sub FW_OptModOrModZ2(ByVal Parameter As CSetFunctionInfo)
    
    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 2 Then
        Err.Raise 9999, "FW_OptEscape", "The number of FW_OptEscape's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Or Flg_Illum_Disable = 1 Then Exit Sub
    '------------------------------------------------------
    
    'Call "OptIni" in case the following object "OptCond" is nothing. (Optcond�I�u�W�F�N�g��Nothing��������OptIni��������)
    Call sub_CheckOptCond
    
'=========Before TestCondition Check Start ======================
    If EEEAUTO_AUTO_MODIFY_TESTCONDITION = True Then
        Dim eMode As eTestCnditionCheck
        eMode = TCC_ILLUMINATOR_MODZ2
        Call CheckBeforeTestCondition(eMode, Parameter)
        If Not IsValidTestCondition(eMode, Parameter) Then Exit Sub
        
    End If
'=========Before TestCondition Check End ========================

    OptCheckCounter = 0
        
    'Init
    'Up   NSIS-5 Escape Point
    'Down
    
    If OptCond.IllumMaker = NIKON Then
        '���݂�F�l�A�������擾
        With Parameter
            If OptCond.IllumModel = "N-SIS3" Or OptCond.IllumModel = "N-SIS5" Or OptCond.IllumModel = "N-SIS5KAI" Then
                '�������̈ړ���ֈړ�
                Call OptModZ_NSIS5(.Arg(1), .Arg(0))
            ElseIf OptCond.IllumModel = "N-SIS3KAI" Then
                'F�l�����̈ړ���ֈړ�
                Call OptMod(.Arg(1), .Arg(0))
            End If
        End With
    End If
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �w�莞�ԕ�Wait������
'
'�p�����[�^:
'    [Arg0]      In   Wait����(s)
'
'�߂�l:
'
'���ӎ���:
'     TheHdw.Wait�Ŗⓚ���p�ɑ҂�
'
Public Sub FW_DebugWait(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    If Parameter.ArgParameterCount() <> 1 Then
        Err.Raise 9999, "FW_DebugWait", "The number of FW_DebugWait's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    
    If Not IsNumeric(Parameter.Arg(0)) Then
        Err.Raise 9999, "FW_DebugWait", "FW_DebugWait's Arg0 is invalid type." & " @ " & Parameter.ConditionName
    End If
    
    Dim dblWaitTime As Double
    dblWaitTime = Parameter.Arg(0)
    
    'Wait����
    Call TheHdw.WAIT(dblWaitTime)
        
    Exit Sub

ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �p�^���̊J�n�������Ȃ�(�I�����܂��Ȃ�)
'
'�p�����[�^:
'    [Arg0]      In   �p�^����
'    [Arg1]      In   TSB�V�[�g��
'    [Arg2]      In   ���s��̃E�F�C�g�^�C��(�ȗ��\ �ȗ���Wait�Ȃ�)
'
'�߂�l:
'
'���ӎ���:IP750 or Decoder Pat�́A��p�Őݒ肷��B
'
Public Sub PatSet(ByVal tmpPatName As String, Optional timeSetName As String = "")

    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
    Const PAT_START_LABEL As String = "START"
            
    Call StopPattern_Halt 'EeeJob�֐�
    Call SetTimeOut 'EeeJob�֐�
    
    '�������W�X�^�Ή����[�`�� Start
    '���W�X�^�ݒ蕔Only(keep_alive)�FPatRun
    '���W�X�^�ݒ�+�쓮��:PatSet
    Dim tmpPatGroupName() As String
    Dim i As Integer
    tmpPatGroupName = Split(tmpPatName, ",")
    
    PatCheckCounter = 0
    
    For i = 0 To UBound(tmpPatGroupName)
        If i < UBound(tmpPatGroupName) Then
            With TheHdw.Digital
                Call .Timing.Load(timeSetName)
                Call .Patterns.pat(tmpPatGroupName(i)).Run(PAT_START_LABEL)
            End With
            If TheExec.RunOptions.AutoAcquire = True Then
                Dim iStatus As Long
                If TheHdw.Digital.Patgen.IsRunningAnySite = True Then      'True:Still Running
                    iStatus = 0
                ElseIf TheHdw.Digital.Patgen.IsRunningAnySite = False Then 'False:haltend or keepalive
                    iStatus = 1
                End If
                
                While (iStatus <> 1)
                    If PatCheckCounter < 999 Then
                        TheHdw.TOPT.Recall
                        PatCheckCounter = PatCheckCounter + 1
                        Call WaitSet(10 * mS)
                        Exit Sub
                    End If
                Wend
            End If
        ElseIf i = UBound(tmpPatGroupName) Then
            With TheHdw.Digital
                Call .Timing.Load(timeSetName)
                Call .Patterns.pat(tmpPatGroupName(i)).Start(PAT_START_LABEL)
            End With
        End If
    Next i
    '�������W�X�^�Ή����[�`�� End
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub

'���e:
'   �p�^�[���o�[�X�g���s���}�N����Call����B
'
'�p�����[�^:
'    [Arg0]      In   �}�N����
'
'�߂�l:
'

Public Sub FW_PatSetCustomMacroA(ByVal Parameter As CSetFunctionInfo)

    On Error GoTo ErrHandler
    
    '---------- SIMULATION ONLY ---------------------------
    If Flg_Simulator = 1 Then Exit Sub
    '------------------------------------------------------
    
    If Parameter.ArgParameterCount <> 3 And Parameter.ArgParameterCount <> 4 Then
         Err.Raise 9999, "FW_PatSetCustomMacroA", "The number of FW_PatSetCustomMacroA's arguments is invalid." & " @ " & Parameter.ConditionName
    End If
    
    Dim strPatGroupName As String
    Dim strTsbName As String
    Dim strMacroName As String
        
    With Parameter
        strPatGroupName = .Arg(0)
        strTsbName = .Arg(1)
        strMacroName = .Arg(2)
    End With
    
    Call Application.Run(strMacroName, strPatGroupName, strTsbName)
    
    Exit Sub
    
ErrHandler:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    
End Sub



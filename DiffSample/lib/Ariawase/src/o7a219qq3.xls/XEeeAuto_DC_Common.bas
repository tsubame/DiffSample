Attribute VB_Name = "XEeeAuto_DC_Common"
'�T�v:
'
'
'�ړI:
'   DC�p���Z�p���W���[��
'
'�쐬��:
'   2011/12/11 Ver0.1 D.Maruyama
'   2011/12/19 Ver0.2 D.Maruyama    Arg20����ĂԂ悤�ɕύX
'                                   �ȉ��̊֐���ǉ�
'                                    ResultSubtract_f
'                                    ResultCalcCommonDifferential_f
'                                    ResultCalcHSImpedanceMismatch_f
'
'   2011/12/21 Ver0.3 D.Maruyama    ResultCalcLPImpedance_f�֐��̒l�̓Ǎ����
'                                   TheDcTest��TheResult�ɕύX
'                                   ResultCalcStpx_f��ǉ�
'
'   2011/12/22 Ver0.4 D.Maruyama    TestInstance�����Arg�̎��o�����֐���
'
'   2012/02/01 Ver0.5 D.Maruyama    DCTestScenario��PreBody�֐���ǉ�
'
'   2012/03/07 Ver0.6 D.Maruyama    CalcOneLSB_f���W���b�W����܂ōs���悤�ɕύX
'   2012/03/17 Ver0.7 D.Maruyama    �ȉ�3�̊֐���ǉ�
'                                   �EResultSTVD_f
'                                   �EResultFPZR_f
'                                   �EResultVCM_f
'   2012/10/19 Ver0.8 K.Tokuyoshi   �ȉ��̊֐���ǉ�
'                                   �EResultCalcLPImpedance_Nega_f
'                                   �EResultDiv_f
'                                   �EResultMultiply_f
'                                   �EResultDiv_Abs_f
'                                   �EResultAbs_Sum_f
'                                   �EResultSubtract_2_f
'                                   �EResultCalcCommonDifferential_f
'                                   �EResultCalcHSImpedance_f
'                                   �EResultCalcLPImpedance_f
'                                   �EResultCalcImpedance_f
'                                   �EResultMin_f
'                                   �EResultCompare_f
'                                   �EResultSubstitution_f
'                                   �EpostDcTestCommonCondition_f
'                                 �@�ȉ��̊֐����C��
'                                   �EResultMax_f
'                                 �@�ȉ��̊֐��𖼑O�ύX
'                                   �EResultAbsDifferetial_f�@�ˁ@ResultSubtract_Abs_f
'                                   �EResultCalcCommonDifferential_f�@�ˁ@ResultCalcCommonDifferential_Abs_f
'                                   �EResultCalcHSImpedance_f�@�ˁ@ResultCalcHSImpedance_Posi_Nega_f
'                                   �EResultCalcLPImpedance_f�@�ˁ@ResultCalcLPImpedance_Posi_f
'                                   �EResultCalcStpx_f�@�ˁ@ResultPixcel_Leak_Ratio_f
'                                   �ECalcOneLSB_f�@�ˁ@ResultCalcOneLSB_f
'                                 �@TheResult.GetResult����mf_GetResult�֕ύX
'   2012/10/22 Ver0.9 K.Tokuyoshi   �ȉ��̊֐���ǉ�
'                                   �EResultBinningFlag_f
'                                   �EResultBinning_f
'                                   �EHold_Voltage_Test_f
'   2012/10/23 Ver1.0 K.Tokuyoshi   �ȉ��̊֐���ǉ�
'                                   �EResultCalcStbyDifferential_f
'                                   �EResultCalcStbyDifferential_Square_f
'                                   �EResultCalcPiezoImpedance_f
'                                   �EResultCalcConductance_f
'                                   �EResultCalcConductance_Posi_Nega_f
'                                   �EResultCalcConductance_Mono_f
'                                   �EResultIndividualCalibrate_f
'   2013/01/31 Ver1.1 K.Hamada       �ȉ��̊֐���ǉ�
'                                   �EResultCalcHS0_HS1Impedance_f
'   2013/02/05 Ver1.2 K.Hamada       �ȉ��̊֐����C��
'                                   �EResultCalcLPImpedance_f
'                                   �EResultMin_f
'   2013/02/06 Ver1.3 K.Hamada       �ȉ��̊֐����C��
'                                   �EResultCalcPiezoImpedance_f
'                                   �EResultCalcConductance_f
'                                   �EResultCalcConductance_Posi_Nega_f
'                                   �EResultCalcConductance_Mono_f
'   2013/02/07 Ver1.4 H.Arikawa      �ȉ��̊֐����C��
'                                   �EResultCalcLPImpedance_Posi_f (���C���f�o�b�OFB)
'                                   �EResultCalcPiezoImpedance_f (DC WG�f�o�b�OFB)
'                                   �EResultCalcConductance_f    (DC WG�f�o�b�OFB)
'                                   �EResultCalcConductance_Posi_Nega_f(DC WG�f�o�b�OFB)
'                                   �EResultCalcConductance_Mono_f  (DC WG�f�o�b�OFB)
'                                   �EResultCalcHS0_HS1Impedance_f  (���C���f�o�b�OFB)
'                                   �EResultCalcOneLsbBasic_f�@�@(�ǉ�)
'   2013/02/07 Ver1.5 H.Arikawa      �ȉ��̊֐����C��
'                           �@      �EsubCurrent_Serial_NoPattern_Test_f
'�@                                 �ESubCurrentTest_NoPattern_GetParameter
'   2013/02/08 Ver1.6 K.Hamada      �ȉ��̊֐����C�� ��Arg�̎w���ύX
'                                   �EResultCalcHSImpedance_Down_f
'                                   �EResultCalcHSImpedance_Up_f
'                                   �EResultCalcLPImpedance_Nega_f
'   2013/02/12 Ver1.7 H.Arikawa     �ȉ��̊֐����C��
'                                   �EResultCalcOneLSB_f
'   2013/02/12 Ver1.8 H.Arikawa     �ȉ��̊֐����C��
'                                   �EResultCalcHS0_HS1Impedance_f
'   2013/02/18 Ver1.9 H.Arikawa     �ȉ��̊֐����C��
'                                   �EResultCalcOneLsbBasic_f
'   2013/02/19 Ver1.A K.Hamada      �ȉ��̊֐���ǉ� ��Arg�̎w���ύX
'                                   �EResultSubtractDiv_f
'                                   �EReturnMaxMinDiff_f
'                                   �EReturnAbsMaxMinValueDCK_f
'   2013/02/25 Ver2.0 H.Arikawa     �ȉ��̊֐���ǉ� ��Arg�̎w���ύX
'                                   �EResultMultiply_f

Option Explicit

Private Const EEE_AUTO_HOLDVOLTAGE_ARGS As Long = 9

'���e:
'   DCTestScenarioFW�̋���PreBody
'
'�p�����[�^:
'[Arg1]         In  �R���f�B�V����1
':
'[ArgN]         In�@�R���f�B�V����N
'
'���ӎ���:
'   �L�q����Ă��鏇�Ԃ�TestCondition���R�[������
'   #EOP��Y��Ȃ�����
Public Function preDcTestCommonCondition_f(argc As Long, argv() As String) As Long

    On Error GoTo ErrorExit

    Call SiteCheck
    Dim i As Long
    
    If argc = 0 Then
        '�G���[�ɂ��Ȃ��Ă悢?
        Exit Function
    End If
        
    '�R���f�B�V���������ԂɎ��{
    For i = 0 To argc - 1
        TheCondition.SetCondition argv(i)
    Next i
    
    preDcTestCommonCondition_f = TL_SUCCESS

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    preDcTestCommonCondition_f = TL_ERROR

End Function

'���e:
'   DCTestScenarioFW�̋���PostBody
'
'�p�����[�^:
'[Arg1]         In  �R���f�B�V����1
':
'[ArgN]         In�@�R���f�B�V����N
'
'���ӎ���:
'   �L�q����Ă��鏇�Ԃ�TestCondition���R�[������
'   #EOP��Y��Ȃ�����
Public Function postDcTestCommonCondition_f(argc As Long, argv() As String) As Long

    On Error GoTo ErrorExit

    Call SiteCheck
    Dim i As Long
    
    If argc = 0 Then
        '�G���[�ɂ��Ȃ��Ă悢?
        Exit Function
    End If
        
    '�R���f�B�V���������ԂɎ��{
    For i = 0 To argc - 1
        TheCondition.SetCondition argv(i)
    Next i
    
    postDcTestCommonCondition_f = TL_SUCCESS

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    postDcTestCommonCondition_f = TL_ERROR

End Function

'��
'��������������������������������������
'IMX145�p���ZVBA�}�N�� :Start
'��������������������������������������
'��


'���e:
'   TestInstance�ɏ����ꂽ�L�[���畽�ς��Ƃ�
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg�擪
':
'[ArgN]         In�@�Ώ�Arg�Ō�
'[ArgN+1]       In  #EOP(End Of Param)
'
'���ӎ���:
'   SUM(Arg1,Arg2,����,ArgN) / N ���v�Z
'   #EOP��Y��Ȃ�����
Public Function ResultAverage_f() As Double

    On Error GoTo ErrorExit

    Call SiteCheck
    
    '�{�֐��ɑ΂���p�����[�^���͕s��B; The number of arguments defined on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔��菬������΃G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultAverage_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂�����; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultAverage_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�������킹
    Dim tmpValue() As Double
    For i = 0 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                retResult(site) = retResult(site) + tmpValue(site)
            End If
        Next site
        Erase tmpValue
    Next i
    
    '���߂�
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = mf_div(retResult(site), lCount)
        End If
    Next site

    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[����ő�l���Ƃ�
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg�擪
':
'[ArgN]         In�@�Ώ�Arg�Ō�
'[ArgN+1]       In  #EOP(End Of Param)
'
'���ӎ���:
'   MAX(Arg1,Arg2,����,ArgN) ���v�Z
'   #EOP��Y��Ȃ�����
' 2012/10/19 K.Tokuyoshi Start�̔�r��ǉ�
Public Function ResultMax_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�{�֐��ɑ΂���p�����[�^���͕s��B; The number of arguments defined on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultMax_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂�����; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultMax_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    
    'MAX�Z�o
    Dim tmpValue0() As Double
    Dim tmpValue1() As Double
    Call mf_GetResult(ArgArr(0), tmpValue0)
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue0(site)
        End If
    Next site
    
    For i = 1 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue1)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If (retResult(site) < tmpValue1(site)) Then retResult(site) = tmpValue1(site)
            End If
        Next site
        Erase tmpValue1
    Next i
    

    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[���獇�v�l���Ƃ�
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg�擪
':
'[ArgN]         In�@�Ώ�Arg�Ō�
'[ArgN+1]       In  #EOP(End Of Param)
'
'���ӎ���:
'   SUM(Arg1,Arg2,����,ArgN) ���v�Z
'   #EOP��Y��Ȃ�����
Public Function ResultSum_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�{�֐��ɑ΂���p�����[�^���͕s��B; The number of arguments defined on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾; To obtain the arguments.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultSum_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂�����; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultSum_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    
    '�������킹
    Dim tmpValue() As Double
    For i = 0 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                retResult(site) = retResult(site) + tmpValue(site)
            End If
        Next site
        Erase tmpValue
    Next i
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[���獷���̐�Βl���Ƃ�
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'
'���ӎ���:
'   Abs(Arg1-Arg2)���v�Z����
'
Public Function ResultSubtract_Abs_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾; To obtain the arguments.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultSubtract_Abs_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultSubtract_Abs_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultSubtract_Abs_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '��Βl����
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = Abs(tmpValue1(site) - tmpValue2(site))
        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[���獷�����Ƃ�
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'
'���ӎ���:
'   Arg1-Arg2���v�Z����
'
Public Function ResultSubtract_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾; To obtain the arguments.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultSubtract_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultSubtract_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultSubtract_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '����
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue1(site) - tmpValue2(site)
        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[���獷���̐�Βl���Ƃ���2�Ŋ���
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'
'���ӎ���:
'   Abs(Arg1-Arg2)/2���v�Z����
'
Public Function ResultCalcCommonDifferential_Abs_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾; To obtain the arguments.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcCommonDifferential_Abs_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcCommonDifferential_Abs_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcCommonDifferential_Abs_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = Abs(tmpValue1(site) - tmpValue2(site)) / 2
        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[����HS�C���s�[�_���X���v�Z����
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'[Arg3]         In  �Ώ�Arg3
'[Arg4]         In  �Ώے萔A
'
'���ӎ���:
'   (Arg1-Arg2)/Arg3/�萔A�@���v�Z����
'
Public Function ResultCalcHSImpedance_Down_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 5
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾; To obtain the arguments.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcHSImpedance_Down_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcHSImpedance_Down_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcHSImpedance_Down_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�C���s�[�_���X�v�Z
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double
    Dim dblCalc1 As Double

    Call mf_GetResult(ArgArr(1), tmpValue1)
    Call mf_GetResult(ArgArr(2), tmpValue2)
    Call mf_GetResult(ArgArr(3), tmpValue3)
    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultCalcHSImpedance_Down_f", "Argument type is Mismatch """ & ArgArr(3) & """ @ " & GetInstanceName
    End If
    dblCalc1 = CDbl(ArgArr(0))

    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult(site) = mf_div(tmpValue3(site), dblCalc1)
            retResult(site) = mf_div((tmpValue2(site) - tmpValue1(site)), Temp_retResult(site))
        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

Public Function ResultCalcHSImpedance_Up_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 5
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾; To obtain the arguments.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcHSImpedance_Up_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcHSImpedance_Up_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcHSImpedance_Up_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�C���s�[�_���X�v�Z
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double
    Dim dblCalc1 As Double

    Call mf_GetResult(ArgArr(1), tmpValue1)
    Call mf_GetResult(ArgArr(2), tmpValue2)
    Call mf_GetResult(ArgArr(3), tmpValue3)
    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultCalcHSImpedance_Up_f", "Argument type is Mismatch """ & ArgArr(3) & """ @ " & GetInstanceName
    End If
    dblCalc1 = CDbl(ArgArr(0))

    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult(site) = mf_div(tmpValue3(site), dblCalc1)
            retResult(site) = mf_div((tmpValue1(site) - tmpValue2(site)), Temp_retResult(site))
        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[���獷���̐�Βl���Ƃ��ĕ��ϒl�ł��
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'
'���ӎ���:
'   Abs(Arg1-Arg2)/Ave(Arg1,Arg2)���v�Z����
'
Public Function ResultCalcHSImpedanceMismatch_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcHSImpedanceMismatch_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcHSImpedanceMismatch_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcHSImpedanceMismatch_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = mf_div(2 * Abs(tmpValue1(site) - tmpValue2(site)), (tmpValue1(site) + tmpValue2(site)))
        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[����LP�C���s�[�_���X���v�Z����
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'[Arg3]         In  �Ώ�Arg3
'[Arg4]         In  �Ώے萔A
'
'���ӎ���:
'   (Arg1-Arg2)/((Arg3-Arg1)/�萔A)���v�Z����
'
Public Function ResultCalcLPImpedance_Nega_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 5
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcLPImpedance_Nega_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcLPImpedance_Nega_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcLPImpedance_Nega_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�C���s�[�_���X�v�Z
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double
    Dim dblCalc1 As Double 'Terminate R Value
    Call mf_GetResult(ArgArr(1), tmpValue1)
    Call mf_GetResult(ArgArr(2), tmpValue2)
    Call mf_GetResult(ArgArr(3), tmpValue3)
    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultCalcLPImpedance_Nega_f", "Argument type is Mismatch """ & ArgArr(3) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(0))
    
    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult(site) = mf_div((tmpValue2(site) - tmpValue3(site)), dblCalc1)
            retResult(site) = mf_div(tmpValue3(site) - tmpValue1(site), Temp_retResult(site))
        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)
    
End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[����LP�C���s�[�_���X���v�Z����
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'[Arg3]         In  �Ώ�Arg3
'[Arg4]         In  �Ώے萔A
'
'���ӎ���:
'   (Arg1-Arg2)/((Arg2-Arg3)/�萔A)���v�Z����
'
Public Function ResultCalcLPImpedance_Posi_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 5
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcLPImpedance_Posi_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcLPImpedance_Posi_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcLPImpedance_Posi_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�C���s�[�_���X�v�Z
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double
    Dim dblCalc1 As Double 'Terminate R Value
    Call mf_GetResult(ArgArr(1), tmpValue1)
    Call mf_GetResult(ArgArr(2), tmpValue2)
    Call mf_GetResult(ArgArr(3), tmpValue3)
    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultCalcLPImpedance_Posi_f", "Argument type is Mismatch """ & ArgArr(3) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(0))
    
    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult(site) = mf_div((tmpValue2(site) - tmpValue3(site)), dblCalc1)
            retResult(site) = mf_div(tmpValue1(site) - tmpValue2(site), Temp_retResult(site))
        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)
    
End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[���珜�Z���s��
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'[Arg3]         In  �Ώ�Arg3
'
'���ӎ���:
'  ( (Arg1+Arg2) / 2   ) / Arg3 ���v�Z����
'
'���ӎ���:
'   #EOP��Y��Ȃ�����

Public Function ResultPixcel_Leak_Ratio_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 4
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔��菬������΃G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultPixcel_Leak_Ratio_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultPixcel_Leak_Ratio_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultPixcel_Leak_Ratio_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�䗦���Z�o
    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    Call mf_GetResult(ArgArr(2), tmpValue3)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult(site) = (tmpValue1(site) + tmpValue2(site)) / 2
            retResult(site) = mf_div(Temp_retResult(site), tmpValue3(site))
        End If
    Next site

    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function
'���e:
'   TestInstance�ɏ����ꂽ�L�[����OneLsb���v�Z���A�o�^����
'
'�p�����[�^:
'[Arg1]         In  LSB�ɂ��邽�߂̌W��
'[Arg2]         In  �e�X�g���ʂ̃L�[
'
'���ӎ���:
'
'
Public Function ResultCalcOneLSB_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult
    
    '�p�����[�^�̎擾
    '�z�萔��菬������΃G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcOneLSB_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcOneLSB_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcOneLSB_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�e�X�g���ʂ̎擾�A�ϐ��̂̂�����
    Dim tmpValue() As Double
    Dim strLSBName As String
    Dim dblMultiply As Double
    Call mf_GetResult(ArgArr(1), tmpValue)
    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultCalcOneLSB_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblMultiply = CDbl(ArgArr(0))
    
    'LSB�̌v�Z
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue(site) * dblMultiply
        End If
    Next site

    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)
    
End Function
'��
'��������������������������������������
'IMX145�p���ZVBA�}�N�� :End
'��������������������������������������
'��

'��
'��������������������������������������
'IMX145�ȊO���ZVBA�}�N�� :Start
'��������������������������������������
'��

'���e:
'   TestInstance�ɏ����ꂽ�L�[���珜�Z���s��
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'
'���ӎ���:
'   (Arg1/Arg2)���v�Z����
'
'���ӎ���:
'   #EOP��Y��Ȃ�����

Public Function ResultDiv_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔��菬������΃G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultDiv_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultDiv_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultDiv_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '����Z���Z�o
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = mf_div(tmpValue1(site), tmpValue2(site))
        End If
    Next site

    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[���珜�Z���s��
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώے萔A
'
'���ӎ���:
'   Arg1 * �萔A ���v�Z����
'
'���ӎ���:
'   #EOP��Y��Ȃ�����

Public Function ResultMultiply_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔��菬������΃G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultMultiply_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultMultiply_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultMultiply_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�|���Z���Z�o
    Dim tmpValue1() As Double
    Dim dblCalc1 As Double
    Call mf_GetResult(ArgArr(1), tmpValue1)
    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultMultiply_f", "Argument type is Mismatch """ & ArgArr(0) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(0))
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue1(site) * dblCalc1
        End If
    Next site

    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[���珜�Z���s��
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώے萔A
'
'���ӎ���:
'   |Arg1| / �萔A���v�Z����
'
'���ӎ���:
'   #EOP��Y��Ȃ�����

Public Function ResultDiv_Abs_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔��菬������΃G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultDiv_Abs_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultDiv_Abs_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultDiv_Abs_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '����Z���Z�o
    Dim tmpValue1() As Double
    Dim dblCalc1 As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    
    If Not IsNumeric(ArgArr(1)) Then
        Err.Raise 9999, "ResultDiv_Abs_f", "Argument type is Mismatch """ & ArgArr(1) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(1))
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = mf_div(Abs(tmpValue1(site)), dblCalc1)
        End If
    Next site

    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[���獇�v�l���Ƃ�
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg�擪
':
'[ArgN]         In�@�Ώ�Arg�Ō�
'[ArgN+1]       In  #EOP(End Of Param)
'
'���ӎ���:
'   Sum(Arg1,Arg2,����,ArgN) ���v�Z
'   #EOP��Y��Ȃ�����
Public Function ResultAbs_Sum_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number of arguments on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔��菬������΃G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_VARIABLE_PARAM) Then
        Err.Raise 9999, "ResultAbs_Sum_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂�����; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultAbs_Sum_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '���v�Z�o
    Dim tmpValue() As Double
    For i = 0 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                retResult(site) = retResult(site) + Abs(tmpValue(site))
            End If
        Next site
        Erase tmpValue
    Next i
    

    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[���獷�����Ƃ�
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'[Arg3]         In  �Ώ�Arg3
'
'���ӎ���:
'   Arg1-Arg2-Arg3���v�Z����
'
Public Function ResultSubtract_2_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 4
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultSubtract_2_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultSubtract_2_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultSubtract_2_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '3�̒l����
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double '2012/11/15 175Debug Arikawa
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    Call mf_GetResult(ArgArr(2), tmpValue3)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue1(site) - tmpValue2(site) - tmpValue3(site)
        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[���獷�����Ƃ���2�Ŋ���
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'
'���ӎ���:
'   (Arg1-Arg2)/2���v�Z����
'
Public Function ResultCalcCommonDifferential_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcCommonDifferential_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcCommonDifferential_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcCommonDifferential_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = (tmpValue1(site) - tmpValue2(site)) / 2
        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[����HS�C���s�[�_���X���v�Z����
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώے萔A
'[Arg3]         In  �Ώے萔B
'
'���ӎ���:
'   �萔A/(Arg1/�萔B) - �萔B�@���v�Z����
'
Public Function ResultCalcHSImpedance_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 4
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcHSImpedance_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcHSImpedance_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcHSImpedance_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�C���s�[�_���X�v�Z
    Dim tmpValue1() As Double
    Dim dblCalc1 As Double
    Dim dblCalc2 As Double

    Call mf_GetResult(ArgArr(0), tmpValue1)
    
    If Not IsNumeric(ArgArr(1)) Then
        Err.Raise 9999, "ResultCalcHSImpedance_f", "Argument type is Mismatch """ & ArgArr(1) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(1))
    
    If Not IsNumeric(ArgArr(2)) Then
        Err.Raise 9999, "ResultCalcHSImpedance_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc2 = CDbl(ArgArr(2))

    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult(site) = mf_div(tmpValue1(site), dblCalc2)
            retResult(site) = mf_div(dblCalc1, Temp_retResult(site)) - dblCalc2
        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[����LP�C���s�[�_���X���v�Z����
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'[Arg3]         In  �Ώے萔A
'
'���ӎ���:
'   |Arg1 - Arg2| / �萔A�@���v�Z����
'
Public Function ResultCalcLPImpedance_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 4
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult
    
    '�p�����[�^�̎擾
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcLPImpedance_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcLPImpedance_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcLPImpedance_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�C���s�[�_���X�v�Z
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim dblCalc1 As Double

    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    If Not IsNumeric(ArgArr(2)) Then
        Err.Raise 9999, "ResultCalcLPImpedance_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(2))

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = mf_div(Abs(tmpValue1(site) - tmpValue2(site)), dblCalc1)
        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[����C���s�[�_���X���v�Z����
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'[Arg3]         In  �Ώے萔A
'
'���ӎ���:
'   ||Arg1| - |Arg2|| / �萔A�@���v�Z����
'
Public Function ResultCalcImpedance_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 4
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcImpedance_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcImpedance_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcImpedance_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�C���s�[�_���X�v�Z
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim dblCalc1 As Double

    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    If Not IsNumeric(ArgArr(2)) Then
        Err.Raise 9999, "ResultCalcImpedance_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(2))

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = mf_div(Abs(Abs(tmpValue1(site)) - Abs(tmpValue2(site))), dblCalc1)
        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[����ő�l���Ƃ�
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg�擪
':
'[ArgN]         In�@�Ώ�Arg�Ō�
'[ArgN+1]       In  #EOP(End Of Param)
'
'���ӎ���:
'   MIN(Arg1,Arg2,����,ArgN) ���v�Z
'   #EOP��Y��Ȃ�����
Public Function ResultMin_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number of arguments on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔��菬������΃G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultMin_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂�����; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultMin_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    
    'MIN�Z�o
    Dim tmpValue() As Double
    Call mf_GetResult(ArgArr(0), tmpValue)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue(site)
        End If
    Next site
    
    For i = 1 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If (retResult(site) > tmpValue(site)) Then retResult(site) = tmpValue(site)
            End If
        Next site
        Erase tmpValue
    Next i
    

    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[���珜�Z���s��
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'
'���ӎ���:
'   Arg1 < Arg2 �̏ꍇ�Ɍ��ʂ�0�����AArg1 > Arg2 �̏ꍇ�Ɍ��ʂ�1������
'
'���ӎ���:
'   #EOP��Y��Ȃ�����

Public Function ResultCompare_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔��菬������΃G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCompare_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCompare_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCompare_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '����Z���Z�o
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If tmpValue1(site) < tmpValue2(site) Then
                retResult(site) = 0
            Else
                retResult(site) = 1
            End If
        End If
    Next site

    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[����l��������
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'
'���ӎ���:
'   ���
'   #EOP��Y��Ȃ�����
Public Function ResultSubstitution_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 2
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔��菬������΃G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultSubstitution_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultSubstitution_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultSubstitution_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '���
    Dim tmpValue1() As Double
    Call mf_GetResult(ArgArr(0), tmpValue1)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue1(site)
        End If
    Next site

    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[����Binning�p��Flag���i�[����
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2�˒萔
'
'���ӎ���:
'   Arg1��Spec���ł���΁y0�z��Spec�O�ł���΁y1�z�ɂ���
'   Arg2��Limit�̐����͈͂��L�ڂ���
'
Public Function ResultBinningFlag_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultBinningFlag_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultBinningFlag_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultBinningFlag_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '����l/Limit�͈�_Get
    Dim tmpValue1() As Double
    Dim dblbinning1 As Double
    Call mf_GetResult(ArgArr(0), retResult)
    
    If Not IsNumeric(ArgArr(1)) Then
        Err.Raise 9999, "ResultBinningFlag_f", "Argument type is Mismatch """ & ArgArr(1) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblbinning1 = CDbl(ArgArr(1))
    
    'Limit_Get
    Dim LoLimit As Double '2012/11/15 175Debug Arikawa
    Dim HiLimit As Double '2012/11/15 175Debug Arikawa
    Call m_GetLimit(LoLimit, HiLimit)
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
        
                Select Case dblbinning1
                    Case 0
                            tmpValue1(site) = 0
                    Case 1
                        If retResult(site) < LoLimit Then
                            tmpValue1(site) = 1
                        End If
                    Case 2
                        If retResult(site) > HiLimit Then
                            tmpValue1(site) = 1
                        End If
                    Case 3
                        If retResult(site) < LoLimit And retResult(site) > HiLimit Then
                            tmpValue1(site) = 1
                        End If
                    Case Else
                End Select

        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)
    
    '���̌��Binning�̌��ʍ��ڂŎg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add("Flg_" & UCase(GetInstanceName), tmpValue1)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[��"Flg_"��t�����ʂ̍��v�l���Ƃ�
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg�擪
':
'[ArgN]         In�@�Ώ�Arg�Ō�
'[ArgN+1]       In  #EOP(End Of Param)
'
'���ӎ���:
'   SUM("Flg_"Arg1,"Flg_"Arg2,����,"Flg_"ArgN) ���v�Z
'   #EOP��Y��Ȃ�����
Public Function ResultBinning_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number of arguments on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔��菬������΃G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultBinning_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂�����; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultBinning_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    
    'Flag�𑫂����킹��
    Dim tmpValue() As Double
    For i = 0 To lCount - 1
        Call mf_GetResult("Flg_" & ArgArr(i), tmpValue)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                retResult(site) = retResult(site) + tmpValue(site)
            End If
        Next site
        Erase tmpValue
    Next i
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[�������v�Z�@���v�Z���A�o�^����
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'[Arg3]         In  �萔A
'[Arg4]         In  �萔B
'
'���ӎ���:
'�@Arg1-(A * (Arg2) + B)���v�Z����
'
Public Function ResultCalcStbyDifferential_f() As Double

    On Error GoTo ErrorExit
        
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 5
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult
    
    '�p�����[�^�̎擾
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcStbyDifferential_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcStbyDifferential_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcStbyDifferential_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�e�X�g���ʂ̎擾�A�ϐ��̂̂�����
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim dblSlope As Double
    Dim dblIntercept As Double
    Call TheResult.GetResult(ArgArr(0), tmpValue1)
    Call TheResult.GetResult(ArgArr(1), tmpValue2)
    
    If Not IsNumeric(ArgArr(2)) Then
        Err.Raise 9999, "ResultCalcStbyDifferential_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblSlope = CDbl(ArgArr(2))
    
    If Not IsNumeric(ArgArr(3)) Then
        Err.Raise 9999, "ResultCalcStbyDifferential_f", "Argument type is Mismatch """ & ArgArr(3) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblIntercept = CDbl(ArgArr(3))
    
    
    'Arg2-(A * (Arg1) + B)�̌v�Z
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue1(site) - ((dblSlope * tmpValue2(site)) + dblIntercept)
        End If
    Next site
    
    '�W���b�W����
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)
    
End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[�������v�Z�A���v�Z���A�o�^����
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'[Arg3]         In  �萔A
'[Arg4]         In  �萔B
'[Arg5]         In  �萔C
'
'���ӎ���:
'�@Arg1-(A * (Arg2)^2 + B * (Arg2) + C)���v�Z����
'
Public Function ResultCalcStbyDifferential_Square_f() As Double

    On Error GoTo ErrorExit
        
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 6
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult
    
    '�p�����[�^�̎擾
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcStbyDifferential_Square_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If

    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcStbyDifferential_Square_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcStbyDifferential_Square_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�e�X�g���ʂ̎擾�A�ϐ��̂̂�����
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim dblSlope1 As Double
    Dim dblSlope2 As Double
    Dim dblIntercept As Double
    Call TheResult.GetResult(ArgArr(0), tmpValue1)
    Call TheResult.GetResult(ArgArr(1), tmpValue2)
    
    If Not IsNumeric(ArgArr(2)) Then
        Err.Raise 9999, "ResultCalcStbyDifferential_Square_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblSlope1 = CDbl(ArgArr(2))
    
    If Not IsNumeric(ArgArr(3)) Then
        Err.Raise 9999, "ResultCalcStbyDifferential_Square_f", "Argument type is Mismatch """ & ArgArr(3) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblSlope2 = CDbl(ArgArr(3))
    
    If Not IsNumeric(ArgArr(4)) Then
        Err.Raise 9999, "ResultCalcStbyDifferential_Square_f", "Argument type is Mismatch """ & ArgArr(4) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblIntercept = CDbl(ArgArr(4))
    
    
    'Arg1-(A * (Arg2)^2 + B * (Arg2) + C)�̌v�Z
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue1(site) - (dblSlope1 * (tmpValue2(site)) ^ 2 + dblSlope2 * tmpValue2(site) + dblIntercept)
        End If
    Next site
    
    '�W���b�W����
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)
    
End Function
'��
'��������������������������������������
'IMX145�ȊO���ZVBA�}�N�� :End
'��������������������������������������
'��

'��
'��������������������������������������
'�ő̒l���� :Start
'��������������������������������������
'��
'���e:
'   TestInstance�ɏ����ꂽ�L�[���獇�v�l���Ƃ�A�������݃��W�X�^�擾
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg�擪
':
'[ArgN]         In�@�Ώ�Arg�Ō�
'[ArgN+1]       In  #EOP(End Of Param)
'
'���ӎ���:
'   SUM(Arg1,Arg2,����,ArgN)���v�Z���A[ParameterTable]�V�[�g���烌�W�X�^���擾����
'   #EOP��Y��Ȃ�����
Public Function ResultIndividualCalibrate_f() As Double

    On Error GoTo ErrorExit

    Call SiteCheck
    
    Dim site As Long
    Dim i As Long
    
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔��菬������΃G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_VARIABLE_PARAM) Then
        Err.Raise 9999, "ResultIndividualCalibrate_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂�����
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultIndividualCalibrate_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�������킹
    Dim tmpValue() As Double
    For i = 0 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                retResult(site) = retResult(site) + tmpValue(site)
            End If
        Next site
        Erase tmpValue
    Next i
    
    '�V�[�g�Ɣ�r���A���W�X�^���擾����
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = ShtParaTable.GetREG(UCase(GetInstanceName), retResult(site))
        End If
    Next

    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function
'��
'��������������������������������������
'�ő̒l���� :End
'��������������������������������������
'��

'��
'��������������������������������������
'Hold_Voltage_Test:Start
'��������������������������������������
'��

Private Function Hold_Voltage_Test_f() As Double

    On Error GoTo ErrorExit

    '����ЂƂłЂƂ̃e�X�g����̂�SiteCheck�͕K�v
    Call SiteCheck
        
    '�ϐ���`
    Dim strResultKey As String              'Arg20�@���ږ�
    Dim strForcePin As String               'Arg21�@�t�H�[�X�[�q
    Dim strMeasurePin As String             'Arg22�@�e�X�g(����)�[�q
    Dim dblStartVoltage As Double           'Arg23�@Start�d��
    Dim dblEndVoltage As Double             'Arg24�@End�d��
    Dim dblStepVoltage As Double            'Arg25�@Step�d��
    Dim dblTargetCurrent As Double          'Arg26�@Target�d��
    Dim strSetParamCondition As String      'Arg27�@����p�����[�^_Opt_�����[
    Dim strPowerCondition As String         'Arg28�@Set_Voltage_�[�q�ݒ�
            
    '����p�����[�^
    Dim lAve As Double                      'TestCondition
    Dim dblClampCurrent As Double           'TestCondition
    Dim dblWait As Double                   'TestCondition
        
    '���ʕϐ�
    Dim retResult(nSite) As Double
            
    '�֐����ϐ�
    Dim dRetValue(nSite) As Double
    Dim resultI(nSite) As Double
    Dim resultID(nSite) As Double
    Dim exitflg(nSite) As Integer
    Dim exitflgJudge As Integer
    Dim TempValue(nSite) As Double
    Dim site As Long
            
    '�ϐ���荞��
    If Not HoldVoltageTest_GetParameter( _
                strResultKey, _
                strForcePin, _
                strMeasurePin, _
                dblStartVoltage, _
                dblEndVoltage, _
                dblStepVoltage, _
                dblTargetCurrent, _
                strSetParamCondition, _
                strPowerCondition) Then
                MsgBox "The Number of Hold_Voltage_Test_f's arguments is invalid!"
                Call DisableAllTest 'EeeJob�֐�
                Exit Function
    End If
            
    '�p�����[�^�ݒ�̊֐����Ă� (FW_SetHoldVoltageParam)
    Call TheCondition.SetCondition(strSetParamCondition)

    lAve = GetHoldVoltageAverageCount(GetInstanceName)
    dblClampCurrent = GetHoldVoltageClampCurrent(GetInstanceName)
    dblWait = GetHoldVoltageWaitTime(GetInstanceName)
    
    'High-HOLD�i���]�j�d�����肩Low-HOLD�i���]�j�d������ŕς�镔���͒ǉ����邩�H�H�H����Ƃ��[�q�ݒ�ɍڂ��Ă��炤���H�H
    
    'HOLD�d������
    Dim i As Long '2012/11/15 175Debug Arikawa
    For i = dblStartVoltage To dblEndVoltage Step dblStepVoltage
        Call SetFVMI(strForcePin, i * V, dblClampCurrent)
        TheHdw.WAIT dblWait * S
        '========== MESURE IO PINS ===============================
        Call MeasureI(strMeasurePin, dRetValue(), lAve)

        For site = 0 To nSite
            If TheExec.sites.site(site).Active And exitflg(site) = 0 Then
                If Flg_Debug = 1 Then TheExec.Datalog.WriteComment strResultKey & strMeasurePin & "  " & i & " " & dRetValue(site)
                If i = dblStartVoltage Then
                    resultI(site) = dRetValue(site)
                Else
                    resultID(site) = resultI(site) - dRetValue(site)
                    If resultID(site) >= 3 * uA Then
                        retResult(site) = i - dblStepVoltage
                        exitflg(site) = 1
                    Else
                        resultI(site) = dRetValue(site)
                    End If
                End If
                exitflgJudge = exitflgJudge + exitflg(site)
            End If
        Next site
        If exitflgJudge > nSite Then Exit For
    Next i

    '����[�q��0V���
    Call SetFVMI(strMeasurePin, 0# * V, dblClampCurrent)
               
    '�W���b�W
    Call test(retResult)

    '�����͕Ԃ���Add����̂�
    Call TheResult.Add(strResultKey, retResult)

    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

Private Function HoldVoltageTest_GetParameter( _
    ByRef strResultKey As String, _
    ByRef strForcePin As String, _
    ByRef strMeasurePin As String, _
    ByRef dblStartVoltage As Double, _
    ByRef dblEndVoltage As Double, _
    ByRef dblStepVoltage As Double, _
    ByRef dblTargetCurrent As Double, _
    ByRef strSetParamCondition As String, _
    ByRef strPowerCondition As String _
    ) As Boolean

    '�ϐ���荞��
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_HOLDVOLTAGE_ARGS) Then
        HoldVoltageTest_GetParameter = False
        Exit Function
    End If
    
On Error GoTo ErrHandler
    strResultKey = ArgArr(0)
    strForcePin = ArgArr(1)
    strMeasurePin = ArgArr(2)
    dblStartVoltage = CDbl(ArgArr(3))
    dblEndVoltage = CDbl(ArgArr(4))
    dblStepVoltage = CDbl(ArgArr(5))
    dblTargetCurrent = CDbl(ArgArr(6))
    strSetParamCondition = ArgArr(7)
    strPowerCondition = ArgArr(8)
On Error GoTo 0

    HoldVoltageTest_GetParameter = True
    Exit Function
    
ErrHandler:

    HoldVoltageTest_GetParameter = False
    Exit Function

End Function
'��
'��������������������������������������
'Hold_Voltage_Test:End
'��������������������������������������
'��

'��
'��������������������������������������
'Piezo�d������(Function):Start
'��������������������������������������
'��

'���e:
'   TestInstance�ɏ����ꂽ�L�[����s�G�]�o�̓C���s�[�_���X���v�Z���A�o�^����
'
'�p�����[�^:
'[Arg0]         In  Arg1
'[Arg1]         In  Arg2
'[Arg2]         In  �萔A
'
'���ӎ���:
'�@�萔A * (1/Arg1- 1/Arg2)���v�Z����
'
Public Function ResultCalcPiezoImpedance_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    Dim site As Long
    
    Dim retResult(nSite) As Double
    Erase retResult
    
    '�p�����[�^�̎擾
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        Err.Raise 9999, "ResultCalcPiezoImpedance_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If

    
    '�e�X�g���ʂ̎擾�A�ϐ��̂̂�����
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim dbldiff As Double
    Call TheResult.GetResult(ArgArr(0), tmpValue1)
    Call TheResult.GetResult(ArgArr(1), tmpValue2)
        
    If Not IsNumeric(ArgArr(2)) Then
        Err.Raise 9999, "ResultCalcPiezoImpedance_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dbldiff = CDbl(ArgArr(2))
        
        
    '�萔A * (1/Arg1- 1/Arg2)�̌v�Z
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = dbldiff * (mf_div(1, tmpValue1(site)) - mf_div(1, tmpValue2(site)))
        End If
    Next site
    
    '�W���b�W����
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)
    
End Function
'��
'��������������������������������������
'Piezo�d������(Function):End
'��������������������������������������
'��

'��
'��������������������������������������
'GCS2 �R���_�N�^���X����(Function):Start
'��������������������������������������
'��

'AFE����ŕK�v�Ȃ�

'���e:
'   TestInstance�ɏ����ꂽ�L�[����o�̓R���_�N�^���X���v�Z���A�o�^����
'
'�p�����[�^:
'[Arg0]         In  Arg1
'[Arg1]         In  �萔A
'
'���ӎ���:
'�@�萔A * (1/Arg1)���v�Z����
'
Public Function ResultCalcConductance_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    Dim site As Long
    
    Dim retResult(nSite) As Double
    Erase retResult
    
    '�p�����[�^�̎擾
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 3) Then
        Err.Raise 9999, "ResultCalcConductance_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If

    
    '�e�X�g���ʂ̎擾�A�ϐ��̂̂�����
    Dim tmpValue1() As Double
    Dim dblCalc1 As Double
    Call TheResult.GetResult(ArgArr(0), tmpValue1)
        
    If Not IsNumeric(ArgArr(1)) Then
        Err.Raise 9999, "ResultCalcConductance_f", "Argument type is Mismatch """ & ArgArr(1) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(1))
        
        
    '�萔A * (1/Arg1)�̌v�Z
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = dblCalc1 * (mf_div(1, tmpValue1(site)))
        End If
    Next site
    
    '�W���b�W����
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)
    
End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[����o�̓R���_�N�^���X(Posi/Nega)���v�Z���A�o�^����
'
'�p�����[�^:
'[Arg0]         In  Arg1
'[Arg1]         In  Arg2
'[Arg2]         In  Arg3
'[Arg3]         In  �萔A(�d����)
'
'���ӎ���:
'�@| (Arg1 - Arg2) / �萔A | - Arg3���v�Z����
'
Public Function ResultCalcConductance_Posi_Nega_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    Dim site As Long
    
    Dim retResult(nSite) As Double
    Erase retResult
    
    '�p�����[�^�̎擾
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 5) Then
        Err.Raise 9999, "ResultCalcConductance_Posi_Nega_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If

    
    '�e�X�g���ʂ̎擾�A�ϐ��̂̂�����
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double
    Dim dblCalc1 As Double
    Call TheResult.GetResult(ArgArr(0), tmpValue1)
    Call TheResult.GetResult(ArgArr(1), tmpValue2)
    Call TheResult.GetResult(ArgArr(2), tmpValue3)
    
    If Not IsNumeric(ArgArr(3)) Then
        Err.Raise 9999, "ResultCalcConductance_Posi_Nega_f", "Argument type is Mismatch """ & ArgArr(4) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(3))
        
    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult
        
    '| (Arg1 - Arg2) / �萔A | - Arg3�̌v�Z
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = Abs(mf_div((tmpValue1(site) - tmpValue2(site)), dblCalc1)) - tmpValue3(site)
        End If
    Next site
    
    '�W���b�W����
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)
    
End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[����o�̓R���_�N�^���X�P�������v�Z���A�o�^����
'
'�p�����[�^:
'[Arg0]         In  Arg1
'[Arg1]         In  Arg2
'[Arg2]         In  �萔A
'
'���ӎ���:
'�@(Arg1 - Arg2) / �萔A���v�Z����
'
Public Function ResultCalcConductance_Mono_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    Dim site As Long
    
    Dim retResult(nSite) As Double
    Erase retResult
    
    '�p�����[�^�̎擾
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        Err.Raise 9999, "ResultCalcConductance_Mono_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If

    
    '�e�X�g���ʂ̎擾�A�ϐ��̂̂�����
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim dblCalc1 As Double
    Call TheResult.GetResult(ArgArr(0), tmpValue1)
    Call TheResult.GetResult(ArgArr(1), tmpValue2)
        
    If Not IsNumeric(ArgArr(2)) Then
        Err.Raise 9999, "ResultCalcConductance_Mono_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(2))
        
    '(Arg1 - Arg2) / �萔A�̌v�Z
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = (mf_div(tmpValue1(site) - tmpValue2(site), dblCalc1))
        End If
    Next site
    
    '�W���b�W����
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)
    
End Function
'��
'��������������������������������������
'GCS2 �R���_�N�^���X����(Function):End
'��������������������������������������
'��
'��
'��������������������������������������
'2013/01/31
'��������������������������������������
'��
'���e:
'   TestInstance�ɏ����ꂽ�L�[����HS�C���s�[�_���X���v�Z����
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'[Arg3]         In  �Ώ�Arg3
'[Arg4]         In  �萔A

'���ӎ���:
'   (Arg1-Arg2)/(Arg3/�萔A�@���v�Z����
'
Public Function ResultCalcHS0_HS1Impedance_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 5
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcHS0_HS1Impedance_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcHS0_HS1Impedance_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcHS0_HS1Impedance_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�C���s�[�_���X�v�Z
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double
    
    Dim dblCalc1 As Double

    Call mf_GetResult(ArgArr(1), tmpValue1)
    Call mf_GetResult(ArgArr(2), tmpValue2)
    Call mf_GetResult(ArgArr(3), tmpValue3)
    

    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultCalcHS0_HS1Impedance_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(0))

    Dim Temp_retResult1(nSite) As Double
    Dim Temp_retResult2(nSite) As Double
    
    Erase Temp_retResult1
    Erase Temp_retResult2
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult1(site) = tmpValue1(site) - tmpValue2(site)
            Temp_retResult2(site) = mf_div(tmpValue3(site), dblCalc1)
            retResult(site) = mf_div(Abs(Temp_retResult1(site)), Temp_retResult2(site))
        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function


'���e:
'   TestInstance�ɏ����ꂽ�L�[����HS�C���s�[�_���X���v�Z����
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'[Arg3]         In  �Ώ�Arg3
'[Arg4]         In  �Ώے萔A
'
'���ӎ���:
'   (Arg1-Arg2)/Arg3/�萔A�@���v�Z����
'
Public Function ResultCalcHSImpedance_Posi_Nega_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    Dim site As Long
    
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        Err.Raise 9999, "ResultCalcHSImpedance_Posi_Nega_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
    End If
    
    '�C���s�[�_���X�v�Z
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim tmpValue3() As Double
    Dim dblCalc1 As Double

    Call mf_GetResult(ArgArr(0), tmpValue1)
    Call mf_GetResult(ArgArr(1), tmpValue2)
    Call mf_GetResult(ArgArr(2), tmpValue3)
    
    If Not IsNumeric(ArgArr(3)) Then
        Err.Raise 9999, "ResultCalcHSImpedance_Posi_Nega_f", "Argument type is Mismatch """ & ArgArr(3) & """ @ " & GetInstanceName
    End If
    dblCalc1 = CDbl(ArgArr(3))

    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult(site) = mf_div(tmpValue3(site), dblCalc1)
            retResult(site) = mf_div((tmpValue1(site) - tmpValue2(site)), Temp_retResult(site))
        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function
'��
'��������������������������������������
'2013/01/31
'��������������������������������������
'��

'���e:
'   TestInstance�ɏ����ꂽ�L�[����OneLsb���v�Z���A�o�^����
'
'�p�����[�^:
'[Arg20]         In  �e�X�g���ʂ̃L�[
'[Arg21]         In  LSB�ɂ��邽�߂̌W��
'[Arg22]             #EOP
'
'���ӎ���:
'
'
Public Function ResultCalcOneLsbBasic_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    'Arg20=DC���荀��
    'Arg21=�W��(�p�����[�^/��)
    'Arg22=#EOP
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 3
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult
    
    '�p�����[�^�̎擾
    '�z�萔��菬������΃G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultCalcOneLSB_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultCalcOneLSB_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultCalcOneLSB_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�e�X�g���ʂ̎擾�A�ϐ��̂̂�����
    Dim tmpValue() As Double
    Dim dblMultiply As Double
    Call mf_GetResult(ArgArr(1), tmpValue)
    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultCalcOneLSB_f", "Argument type is Mismatch """ & ArgArr(2) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblMultiply = CDbl(ArgArr(0))
    
    'LSB�̌v�Z
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = tmpValue(site) * dblMultiply
        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)
    
End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[���獷�����Ƃ�A�W���Ŋ���
'
'�p�����[�^:
'[Arg0]         In  �Ώ�Arg0 ���W��
'[Arg1]         In  �Ώ�Arg1
'[Arg2]         In  �Ώ�Arg2
'
'���ӎ���:
'   (Arg1-Arg2)/Arg0���v�Z����
'
Public Function ResultSubtractDiv_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�p�����[�^��(������"#EOP"���܂�); Number Of Parameters including "#EOP" at the end.
    Const NOF_INSTANCE_ARGS As Long = 4
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultSubtractDiv_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂��A�����ɂ��Ȃ��̂�����
    ' ; Error will be raised in case "#EOP" is absent at the end.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultSubtractDiv_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    ElseIf lCount + 1 <> NOF_INSTANCE_ARGS Then
        Err.Raise 9999, "ResultSubtractDiv_f", """#EOP"" is at illegal position [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    Dim tmpValue1() As Double
    Dim tmpValue2() As Double
    Dim dblCalc1 As Double '�W��
    Call mf_GetResult(ArgArr(1), tmpValue1)
    Call mf_GetResult(ArgArr(2), tmpValue2)

    
    If Not IsNumeric(ArgArr(0)) Then
        Err.Raise 9999, "ResultSubtractDiv_f", "Argument type is Mismatch """ & ArgArr(3) & """ @ " & GetInstanceName
        GoTo ErrorExit
    End If
    dblCalc1 = CDbl(ArgArr(0))
    
    Dim Temp_retResult(nSite) As Double
    Erase Temp_retResult
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Temp_retResult(site) = (tmpValue1(site) - tmpValue2(site))
            retResult(site) = mf_div(Temp_retResult(site), dblCalc1)
        End If
    Next site
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[����ő�l�ƍŏ��l���Ƃ荷�����Ƃ�
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg�擪
':
'[ArgN]         In�@�Ώ�Arg�Ō�
'[ArgN+1]       In  #EOP(End Of Param)
'
Public Function ReturnMaxMinDiff_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�{�֐��ɑ΂���p�����[�^���͕s��B; The number of arguments defined on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ReturnMaxMinDiff_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂�����; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ReturnMaxMinDiff_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    
    'MAX�Z�o
    Dim tmpValue0() As Double
    Dim tmpValue1() As Double
    Dim TempMaxValue(nSite) As Double
    Dim TempMinValue(nSite) As Double
    
    Call mf_GetResult(ArgArr(0), tmpValue0)
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            TempMaxValue(site) = tmpValue0(site)
            TempMinValue(site) = tmpValue0(site)
        End If
    Next site
    
    For i = 1 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue1)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If (TempMaxValue(site) < tmpValue1(site)) Then TempMaxValue(site) = tmpValue1(site)
                If (TempMinValue(site) > tmpValue1(site)) Then TempMinValue(site) = tmpValue1(site)
            End If
        Next site
        Erase tmpValue1
    Next i
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            retResult(site) = TempMaxValue(site) - TempMinValue(site)
        End If
    Next site

    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'���e:
'   TestInstance�ɏ����ꂽ�L�[����DCK��ōő�l�ƍŏ��l���Ƃ��Βl�̑傫���l���Ƃ�
'
'�p�����[�^:
'[Arg1]         In  �Ώ�Arg�擪
':
'[ArgN]         In�@�Ώ�Arg�Ō�
'[ArgN+1]       In  #EOP(End Of Param)
Public Function ReturnAbsMaxMinValueDCK_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�{�֐��ɑ΂���p�����[�^���͕s��B; The number of arguments defined on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ReturnAbsMaxMinValueDCK_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂�����; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ReturnAbsMaxMinValueDCK_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    
    'MAX�Z�o
    Dim tmpValue0() As Double
    Dim tmpValue1() As Double
    Dim TempMaxValue(nSite) As Double
    Dim TempMinValue(nSite) As Double
    Dim DCKValue(nSite) As Double
    
    Call mf_GetResult(ArgArr(0), tmpValue0)
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            DCKValue(site) = tmpValue0(site)
        End If
    Next site
    
    Call mf_GetResult(ArgArr(1), tmpValue1)

    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            TempMaxValue(site) = tmpValue1(site) - DCKValue(site)
            TempMinValue(site) = tmpValue1(site) - DCKValue(site)
        End If
    Next site
    
    
    For i = 1 + 1 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue1)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                If (TempMaxValue(site) < (tmpValue1(site) - DCKValue(site))) Then TempMaxValue(site) = tmpValue1(site) - DCKValue(site)
                If (TempMinValue(site) > (tmpValue1(site) - DCKValue(site))) Then TempMinValue(site) = tmpValue1(site) - DCKValue(site)
            End If
        Next site
        Erase tmpValue1
    Next i
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            If (Abs(TempMaxValue(site)) > Abs(TempMinValue(site))) Then retResult(site) = TempMaxValue(site)
            If (Abs(TempMaxValue(site)) < Abs(TempMinValue(site))) Then retResult(site) = TempMinValue(site)
        End If
    Next site

    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function


'���e:
'   ����d�͂����߂�}�N���B�w�肳���d������e�X�g���x���ƁA�Ή�����d���l�̒l��p����B
'
'�p�����[�^:
'[Arg20]        In  Arg21�Ȍ�̃J�����Ɏw�肳���d������e�X�g���x�����X�g�Ɠ����́u�d���l�v��
'                   ","(�J���})��؂�ŋL�ڂ���B�킴�킴�d���l���w�肷��̂́A�d�����肵�Ă�
'                   �e�X�g�����Ŏg�p���Ă���d���l�ł͂Ȃ��l�Ōv�Z���������Ƃ����A���i�d�l����
'                   �v�������邽�߁B
'[Arg21]...     In�@�d������������Ƃ��̃e�X�g���x�����B
Public Function ReturnPowerConsumption_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�{�֐��ɑ΂���p�����[�^���͕s��B; The number of arguments defined on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    
    '�p�����[�^�̎擾
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ReturnPowerConsumption_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂�����; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ReturnPowerConsumption_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�ŏ��̒l(Arg20)���d���l�B
    Dim strVddValues() As String
    strVddValues = Split(ArgArr(0), ",")
    
    '�e�X�g���x�����Ƃ̐��̈�v���m�F�B
    If UBound(strVddValues) <> lCount - 2 Then
        Err.Raise 9999, "ReturnPowerConsumption_f", "The number of test labels and vdd values do not match."
    End If
    
    '���������_�l�ɕϊ�
    Dim dblVddValues() As Double
    ReDim dblVddValues(UBound(strVddValues))
    For i = 0 To UBound(strVddValues)
        If Not IsNumeric(strVddValues(i)) Then
            Err.Raise 9999, "ReturnPowerConsumption_f", "VDD value list (Arg20) must be comma separated numeric values"
        Else
            dblVddValues(i) = CDbl(strVddValues(i))
        End If
    Next i
    
    '����d�͎Z�o
    Dim tmpIddValue() As Double
    Dim retResult(nSite) As Double
    Erase retResult
    For i = 0 To UBound(dblVddValues)
        Call mf_GetResult(ArgArr(i + 1), tmpIddValue)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active Then
                retResult(site) = retResult(site) + tmpIddValue(site) * dblVddValues(i)
            End If
        Next site
    Next i
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

'��
'��������������������������������������
'�唻_ScrnWait_VBA�}�N�� :Start
'��������������������������������������
'��

'���e:
'   �X�N���[�j���O�����s��������Ԃ��e�X�g���ڂ֕Ԃ��B
'
'�p�����[�^:
'
'���ӎ���:
'

Public Function ScreeninApplyWait_f() As Double

    On Error GoTo ErrorExit

    '�ϐ���`
    'Arg 20: �������
    'Arg 21: ����v���d�l���Ɏw�肳���X�N���[�j���O��Wait����
    'Arg 22: �e�X�g���x��
    Dim strSetCondition As String       'Arg20: [Test Condition]'s condition name excluding screening wait.
    Dim dblWaitTime As Double           'Arg21�@����d��
    Dim strResultKey As String          'Arg22�@���ږ�
    
    If Flg_Scrn = 1 And Flg_Tenken = 0 Then

        '�ϐ���荞��
        If Not ScreeningWait_GetParameter( _
                    strSetCondition, _
                    dblWaitTime, _
                    strResultKey) Then
                    
                    MsgBox "The Number of ScreeninApplyWait_f's arguments is invalid!"
                    Call DisableAllTest 'EeeJob�֐�
                    Exit Function
                    
        End If
            
        '========== Set Condition ===============================
        Call TheCondition.SetCondition(strSetCondition)
        
        TheHdw.WAIT (dblWaitTime)    '�������
    Else
        Exit Function
    End If
        
    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�

End Function

Private Function ScreeningWait_GetParameter(ByRef strSetCondition As String, _
                                            ByRef dblWaitTime As Double, _
                                            ByRef strResultKey As String) As Boolean

    '�ϐ���荞��
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        ScreeningWait_GetParameter = False
        Exit Function
    End If
    
    Dim tempstr As String      '�ꎞ�ۑ��ϐ�
    Dim tempArrstr() As String '�ꎞ�ۑ��z��
    
On Error GoTo ErrHandler
    strSetCondition = ArgArr(0)        'Arg24-1: [Test Condition]'s condition name for common environment setup.
    dblWaitTime = ArgArr(1)       'Arg22: Force voltage (PPS value) for test pin.
    strResultKey = ArgArr(2)                'Arg20: Test label name.
On Error GoTo 0

    ScreeningWait_GetParameter = True
    Exit Function
    
ErrHandler:

    ScreeningWait_GetParameter = False
    Exit Function

End Function


'���e:
'   TestInstance�ɏ����ꂽ�L�[�ƁA�e�L�[�ɑ΂���d�ݕt���W������W�����������l�̑��a���Ƃ�B
'       (To obtain the sum of previous test results multiplied with user factors)
'�����F
'   IMX227�ł�MIPI�̎����쎞�̑z�����d������̂��߂ɊJ���BMIPI�̎�����ł́A1H���Ԃɑ΂���
'   ���Ԕ�Ƃ��āA��LP���[�h�쓮��65%�AHS���[�h�쓮��35%���߂�B���ꂼ��̓��쎞�̏���d���ʂ�
'   ��ɑ��肵�Ă����A���̒l��0.35�Ȃ�т�0.65�������Ęa���Ƃ邱�Ƃ�ړI�Ƃ����B
'     ���̂悤�Ȍv�Z���������ėp���������̂��{�֐��ł���B
'       (Originally it is intended to calculate reliable current consumption under MIPI burst mode.
'        In the test, both current consumption value under MIPI-HS and LP burst is measured
'        respectively, and then total cunsumption value is calculated by the following equation
'               [Total current] = 0.35 * [current under HS-Burst] + 0.65 * [current under LP-burst]
'       where, 0.35 and 0.65 is the ratio of the burst period in 1H
'
'
'�p�����[�^:
'[Arg20]        In  Arg21����"#EOP"�R�[�h�̂���Argument��-1��܂łɕ���ł���
'                   �L�[(�e�X�g���x����)�Ɠ����́u�d�ݕt���W���v���A","
'                   (���p�J���})��؂�ŗ񋓂��ꂽ���́B����v���d�l����
'                   �u��/�p�����[�^�v�ɂ�����"$�W��"���̉E�ӁB
'                       ��) 0.35,0.65
'                   (Comma separated weight factor values. The number of values must equal to
'                   the number of test label names specified at [ArgN].)
'[ArgN]         In�@(N��21�ȏ�̐���)�ΏۃL�[(�e�X�g���x����)
'                   (Target test label names.)
'[ArgN+1]       In  #EOP(End Of Param)
'
'���ӎ���:
'   #EOP��Y��Ȃ�����
Public Function ResultWeightFactorSum_f() As Double

    On Error GoTo ErrorExit
    
    Call SiteCheck
    
    '�{�֐��ɑ΂���p�����[�^���͕s��B; The number of arguments defined on the Test Instances sheet is variable.
    Const NOF_INSTANCE_ARGS As Long = EEE_AUTO_VARIABLE_PARAM
    Dim site As Long
    Dim i As Long
    Dim retResult(nSite) As Double
    Erase retResult

    '�p�����[�^�̎擾; To obtain the arguments.
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, NOF_INSTANCE_ARGS) Then
        Err.Raise 9999, "ResultSum_f", "The Number of arguments is invalid! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�I�������񂪌�����Ȃ��̂�����; To check the presence of "#EOP" and determine the number of arguments.
    Dim IsFound As Boolean
    Dim lCount As Long
    IsFound = False
    For i = 0 To UBound(ArgArr)
        If (ArgArr(i) = "#EOP") Then
            lCount = i '0�n�܂�Ȃ̂�#EOP�̈ʒu���L�������̐��ƂȂ�
            IsFound = True
            Exit For
        End If
    Next
    If Not IsFound Then
        Err.Raise 9999, "ResultWeightFactorSum_f", """#EOP"" is not found! [" & GetInstanceName & "] !"
        GoTo ErrorExit
    End If
    
    '�������̏d�ݕt���W���̕���
    Dim strWeightFactors() As String
    strWeightFactors = Split(ArgArr(0), ",")
    If UBound(strWeightFactors) <> lCount - 2 Then
        Call MsgBox("Error occurred! : The number of test keys and factors do not match.")
        Call DisableAllTest
        Call test(retResult)
        Exit Function
    End If
    Dim dblWeightFactors() As Double
    ReDim dblWeightFactors(UBound(strWeightFactors))
On Error GoTo NotNumericError
    For i = 0 To UBound(strWeightFactors)
        dblWeightFactors(i) = CDbl(strWeightFactors(i))
    Next i
On Error GoTo ErrorExit
    
    
    '�������킹
    Dim tmpValue() As Double
    For i = 1 To lCount - 1
        Call mf_GetResult(ArgArr(i), tmpValue)
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                retResult(site) = retResult(site) + tmpValue(site) * dblWeightFactors(i - 1)
            End If
        Next site
        Erase tmpValue
    Next i
    
    '�W���b�W
    Call test(retResult)
    
    '���̌�̃e�X�g�Ŏg�p�ł���悤��ResultManager�ɓo�^���Ă���
    Call TheResult.Add(UCase(GetInstanceName), retResult)

    Exit Function
    
NotNumericError:
    MsgBox "Error Occurred !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description & vbCrLf & "Weight Factors (Arg20 of Test Instances sheet) must be comma separated numeric values."
    Call DisableAllTest
    Call test(retResult)
    Exit Function
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)

End Function

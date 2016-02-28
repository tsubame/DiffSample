Attribute VB_Name = "XEeeAuto_DC_SubCurrent"
'�T�v:
'
'
'�ړI:
'   STB�d��������s�����߂̃��W���[��
'
'�쐬��:
'   2011/12/06 Ver0.1 D.Maruyama
'   2011/12/06 Ver0.2 D.Maruyama �s�����\�[�X����ǉ�
'   2011/12/16 Ver0.3 D.Maruyama TestInstance��Arg�J�n�ʒu��ύX
'   2011/12/22 Ver0.4 D.Maruyama TestInstance�����Arg�̎��o�����֐���
'   2012/01/23 Ver0.5 D.Maruyama �p�����[�^�̈ꕔ��TestCondition�V�[�g�o�R�ɕύX
'   2012/02/03 Ver0.6 D.Maruyama Key����TestInstance�o�R�Ŏ擾����悤�ɂɕύX
'   2012/02/14 Ver0.7 D.Maruyama SUB�d������̃p�����[�^�ݒ�����̃C���X�^���X��
'                                TestCondition�V�[�g����Ăׂ�悤�ɕύX�B
'   2012/02/20 Ver0.8 D.Maruyama �W���b�W�͕ʂɂ��Ȃ��Ƃ����Ȃ��̂ŁAResultManager��Add����L�[����ʂɓn��
'   2012/03/16 Ver0.9 D.Maruyama APMUUB�̐؂藣����CUB�ɂȂ��Ă����̂��C��
'   2012/10/19 Ver1.2 K.Tokuyoshi �啝�ɏC��
'   2012/12/26 Ver1.3 H.Arikawa  GndSeparateBySite��Private�֐�����Public�֐��֕ύX
'   2013/01/22 Ver1.4 H.Arikawa  subCurrent_Test_f��SubCurrentTestIfNeeded_f�֕ύX�A�p�����[�^�ݒ�C���B
'   2013/01/25 Ver1.5 H.Arikawa  �C���B
'   2013/01/31 Ver1.6 H.Arikawa  �C���B
'   2013/02/05 Ver1.7 H.Arikawa  Debug���e���f�B
'   2013/02/07 Ver1.8 H.Arikawa  subCurrent_Serial_Test_f���C���B
'   2013/02/07 Ver1.9 H.Arikawa  SubCurrentTest_NoPattern_GetParameter�AsubCurrent_Serial_NoPattern_Test_f��ǉ��B
'   2013/02/12 Ver2.0 H.Arikawa  Arg����`�����C���BsubCurrent_Serial_NoPattern_Test_f�C���B
'   2013/02/22 Ver2.1 H.Arikawa  subCurrent_Serial_NoPattern_Test_f�C���B
'   2013/03/11 Ver2.2 K.Hamada   SubCurrentTestIfNeeded_f �C��
'                                SubCurrentNonScenario_Measure_f�ǉ�
'                                subCurrentNonScenarioSeriParaJudge_f�ǉ�


Option Explicit

'�p�����������V���A��������s�Ȃ��ۂ̃V���A������p�̃p�����[�^���B������"#EOP"���܂ށB
'   Number of arguments for serial current measurement following parallel measurement and its judge.
'   "#EOP" at the end is also accounted.
Public Const EEE_AUTO_SERIPARA_SERI_ARGS As Long = 4
Public Const EEE_AUTO_BPMU_PARA_ARGS As Long = 4

'�V���A������݂̂̓d������p�̃p�����[�^���B������"#EOP"���܂ށB
'   Number of arguments for serial current measurement only test.
'   "#EOP" at the end is also accounted.
Public Const EEE_AUTO_SERIAL_ARGS As Long = 7

'�p�����������V���A������ɓ˓����邩�ǂ����̔���p�̃p�����[�^���B������"#EOP"���܂�
'   Number of arguments for judgement of execution for serial current measurement including "#EOP" at the end.
Public Const EEE_AUTO_SUB_SERIPARA_JUDGE_ARGS As Long = 5

'�V���A������݂̂̓d������p�̃p�����[�^���B������"#EOP"���܂ށB(�p�^�[�������p)
Public Const EEE_AUTO_SUBCURRENT_ARGS As Long = 6

Public Go_Serial_Mesure As Boolean

'��
'��������������������������������������
'subCurrent_Serial_Test�p:Start
'��������������������������������������
'��

Private Function subCurrent_Serial_Test_f() As Double

    On Error GoTo ErrorExit

    '����ЂƂłЂƂ̃e�X�g����̂�SiteCheck�͕K�v
    Call SiteCheck
        
    '�ϐ���`
    Dim strResultKey As String              'Arg20�@���ږ�
    Dim strPin As String                    'Arg21�@�e�X�g�[�q
    Dim dblForceVoltage As Double           'Arg22�@����d��
    Dim strSetParamCondition As String      'Arg23�@����p�����[�^_Opt_�����[
    Dim strPowerCondition As String         'Arg24�@Set_Voltage_�[�q�ݒ�
    Dim strPatternCondition As String       'Arg25�@Pattern
            
    '����p�����[�^
    Dim lAve As Double                      'TestCondition
    Dim dblClampCurrent As Double           'TestCondition
    Dim dblWait As Double                   'TestCondition
    Dim dblPinResourceName As String        'TestCondition
        
    '���ʕϐ�
    Dim retResult(nSite) As Double
            
    '�֐����ϐ�
    Dim Flg_Active(nSite) As Long
    Dim TempValue(nSite) As Double
    Dim site As Long
            
    '�ϐ���荞��
    If Not SubCurrentTest_GetParameter( _
                strResultKey, _
                strPin, _
                dblForceVoltage, _
                strSetParamCondition, _
                strPowerCondition, _
                strPatternCondition) Then
                MsgBox "The Number of subCurrent_Serial_Test_f's arguments is invalid!"
                Call DisableAllTest 'EeeJob�֐�
                Exit Function
    End If
            
    '�p�����[�^�ݒ�̊֐����Ă� (FW_SetSubCurrentParam)
    Call TheCondition.SetCondition(strSetParamCondition)

    lAve = GetSubCurrentAverageCount(GetInstanceName)
    dblClampCurrent = GetSubCurrentClampCurrent(GetInstanceName)
    dblWait = GetSubCurrentWaitTime(GetInstanceName)
    dblPinResourceName = GetSubCurrentPinResourceName(GetInstanceName)

    'Active�T�C�g�̊m�F
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Flg_Active(site) = 1
        End If
    Next site

    'SUB�d������̊m�F
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            '========== ���g�psite��BetaGND�؂藣�� ===============================
            Call GndSeparateBySite(site)
            
            '========== Set Condition ===============================
            Call TheCondition.SetCondition(strPowerCondition)

             '========== Force Voltage ===============================
            If StrComp(dblPinResourceName, "BPMU", vbTextCompare) = 0 Then
                Call SetFVMI_BPMU(strPin, 0 * V, dblClampCurrent, site)
                Call DisconnectPins(strPin, site)
                Call SetFVMI_BPMU(strPin, dblForceVoltage, dblClampCurrent, site)
            Else
                Call SetFVMI(strPin, 0 * V, dblClampCurrent, site)
                Call DisconnectPins(strPin, site)
                Call SetFVMI(strPin, dblForceVoltage, dblClampCurrent, site)
            End If
            
             '========== Set Pattern ===============================
            Call TheCondition.SetCondition(strPatternCondition)
            TheHdw.WAIT dblWait
            
             '========== Measure Current ===============================
            If StrComp(dblPinResourceName, "BPMU", vbTextCompare) = 0 Then
                Call MeasureI_BPMU(strPin, retResult, lAve, site)
                Call SetFVMI_BPMU(strPin, 0 * V, dblClampCurrent, site)
                Call DisconnectPins(strPin, site)
            Else
                Call MeasureI(strPin, retResult, lAve, site)
                Call SetFVMI(strPin, 0 * V, dblClampCurrent, site)
                Call DisconnectPins(strPin, site)
            End If

             '========== ���g�psite��BetaGND�߂� =======================
            Call GndConectBySite(site, Flg_Active)
        End If
    Next site
    
    '�p�^�[����~
    Call StopPattern 'EeeJob�֐�
  
    'All_Open�y��Disconnect����������
    Call PowerDownAndDisconnect
           
    '�����͕Ԃ���Add����̂�
    Call test(retResult)
    Call TheResult.Add(strResultKey, retResult)

    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)
    Call TheResult.Add(strResultKey, retResult)

End Function
'��
'��������������������������������������
'subCurrent_Serial_Test�p:End
'��������������������������������������
'��
'2013/01/22 H.Arikawa Arg23,24,25 Get -> Arg24 �J���}��؂�Ή�
Private Function SubCurrentTest_GetParameter( _
    ByRef strResultKey As String, _
    ByRef strPin As String, _
    ByRef dblForceVoltage As Double, _
    ByRef strSetParamCondition As String, _
    ByRef strPowerCondition As String, _
    ByRef strPatternCondition As String _
    ) As Boolean

    '�ϐ���荞��
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_SERIAL_ARGS) Then
        SubCurrentTest_GetParameter = False
        Exit Function
    End If
    
On Error GoTo ErrHandler
    strResultKey = ArgArr(0)                'Arg20: Test label name.
    strPin = ArgArr(1)                      'Arg21: Test pin name
    dblForceVoltage = CDbl(ArgArr(2))       'Arg22: Force voltage (PPS value) for test pin.
    strSetParamCondition = ArgArr(3)        'Arg23: [Test Condition]'s condition name for common environment setup.
    strPowerCondition = ArgArr(4)           'Arg24: [Test Condition]'s condition name for device setup.
    strPatternCondition = ArgArr(5)         'Arg25: [Test Condition]'s condition name for patttern burst.
On Error GoTo 0

    SubCurrentTest_GetParameter = True
    Exit Function
    
ErrHandler:

    SubCurrentTest_GetParameter = False
    Exit Function

End Function
'��
'��������������������������������������
'subCurrent_Serial_Test�p:End
'��������������������������������������
'��

'��
'��������������������������������������
'subCurrent_Serial_Test�p:End
'��������������������������������������
'��
Private Function getParam_SerialMeasureAfterParallel( _
    ByRef strPin As String, _
    ByRef dblForceVoltage As Double, _
    ByRef strSetParamCondition As String, _
    ByRef strPowerCondition As String, _
    ByRef strPatternCondition As String _
    ) As Boolean

    '�ϐ���荞��
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_SERIPARA_SERI_ARGS) Then
        getParam_SerialMeasureAfterParallel = False
        Exit Function
    End If
    
    Dim tempstr As String      '�ꎞ�ۑ��ϐ�
    Dim tempArrstr() As String '�ꎞ�ۑ��z��
    
On Error GoTo ErrHandler
    strPin = ArgArr(0)                      'Arg20: Test pin name
    dblForceVoltage = CDbl(ArgArr(1))       'Arg21: Force voltage (PPS value) for test pin.
    tempstr = ArgArr(2)
    tempArrstr = Split(tempstr, ",")
    strSetParamCondition = tempArrstr(0)        'Arg22-1: [Test Condition]'s condition name for common environment setup.
    strPowerCondition = tempArrstr(1)           'Arg22-2: [Test Condition]'s condition name for device setup.
    strPatternCondition = tempArrstr(2)         'Arg22-3: [Test Condition]'s condition name for patttern burst.
On Error GoTo 0

    getParam_SerialMeasureAfterParallel = True
    Exit Function
    
ErrHandler:

    getParam_SerialMeasureAfterParallel = False
    Exit Function

End Function


'��
'��������������������������������������
'subCurrent_Serial_Test�p:End
'��������������������������������������
'��
Private Function getParam_SerialMeasureAfterParallel_NoPattern( _
    ByRef strPin As String, _
    ByRef dblForceVoltage As Double, _
    ByRef strSetParamCondition As String, _
    ByRef strPowerCondition As String _
    ) As Boolean

    '�ϐ���荞��
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_SERIPARA_SERI_ARGS) Then
        getParam_SerialMeasureAfterParallel_NoPattern = False
        Exit Function
    End If
    
    Dim tempstr As String      '�ꎞ�ۑ��ϐ�
    Dim tempArrstr() As String '�ꎞ�ۑ��z��
    
On Error GoTo ErrHandler
    strPin = ArgArr(0)                      'Arg20: Test pin name
    dblForceVoltage = CDbl(ArgArr(1))       'Arg21: Force voltage (PPS value) for test pin.
    tempstr = ArgArr(2)
    tempArrstr = Split(tempstr, ",")
    strSetParamCondition = tempArrstr(0)        'Arg22-1: [Test Condition]'s condition name for common environment setup.
    strPowerCondition = tempArrstr(1)           'Arg22-2: [Test Condition]'s condition name for device setup.
On Error GoTo 0

    getParam_SerialMeasureAfterParallel_NoPattern = True
    Exit Function
    
ErrHandler:

    getParam_SerialMeasureAfterParallel_NoPattern = False
    Exit Function

End Function
'��
'��������������������������������������
'subCurrent_Serial_Test�p:End
'��������������������������������������
'��
Private Sub GndSeparateBySite(ByVal targetSite As Long)

    Dim site As Long

    '========== �SSITE����ACTIVE�� (���sSITE�ȊO) =============================
    For site = 0 To nSite
        If site <> targetSite Then TheExec.sites.site(site).Active = True
    Next site
    
    '========== �����O�̃f�o�C�X��~�ݒ� ======================================
    '�p�^�[����~
    Call StopPattern 'EeeJob�֐�
          
    'All_Open�y��Disconnect����������
    Call PowerDownAndDisconnect
    
    '========== ���g�pSITE��GND���� ===========================================
    For site = 0 To nSite
        If site <> targetSite Then Call SET_RELAY_CONDITION("GND_Separate_Site" & CStr(site), "-") '2012/11/16 175Debug Arikawa
    Next site
                  
    '========== ���g�pSITE�̒�~ ==============================================
    For site = 0 To nSite
        If site <> targetSite Then TheExec.sites.site(site).Active = False
    Next site

End Sub

Private Sub GndConectBySite(ByVal targetSite As Long, ByRef ActiveSiteFlg() As Long)

    Dim site As Long

    '========== ������SITE��GND =============================
    For site = 0 To nSite
        If site <> targetSite Then
            If ActiveSiteFlg(site) = 1 Then
                TheExec.sites.site(site).Active = True
                Call SET_RELAY_CONDITION("GND_Beta_Site" & CStr(site), "-") '2012/11/16 175Debug Arikawa
            End If
        End If
    Next site

End Sub

'��
'��������������������������������������
'subCurrent_Parallel��Serial_Judge�p:Start
'��������������������������������������
'��

Private Function subCurrentSeriParaJudge_f() As Double

    '����ЂƂłЂƂ̃e�X�g����̂�SiteCheck�͕K�v
    Call SiteCheck
    
    Go_Serial_Mesure = False
    
    '�ϐ���`
    Dim strResultKey As String              'Arg20�@Test label
    Dim dblLoJudgeLimit As Double           'Arg21�@Serial�˓�Low���~�b�g
    Dim dblHiJudgeLimit As Double           'Arg22�@Serial�˓�High���~�b�g
    Dim dblHiLoLimValid As Long             'Arg23�@Serial�˓����~�b�g�̗L���͈�
    
    '�ϐ���荞��
    If Not Sub_SeriParaJudge_GetParameter( _
                strResultKey, _
                dblLoJudgeLimit, _
                dblHiJudgeLimit, _
                dblHiLoLimValid) Then
                MsgBox "The Number of subCurrentSeriParaJudge_f's arguments is invalid!"
                Call DisableAllTest 'EeeJob�֐�
                Exit Function
    End If
    
    'SeriParaJudge
    Call mf_Sub_SeriParaJudge(strResultKey, dblLoJudgeLimit, dblHiJudgeLimit, dblHiLoLimValid, Go_Serial_Mesure)
    
End Function


'��
'��������������������������������������
'subCurrent_Parallel��Serial_Test�p:Start(�p���������葤)
'��������������������������������������
'��
'2013/12/05 T.Morimoto
'               �_�R���񂪁AIMX219�ɂ����Ēǉ����ڂƂ���HW�X�^���o�C�ł̃V���A������
'               (�p����������㔻�聨�V���A������)���������ꂽ���Ƃɔ����ǉ��B
'               �@BPMU�Ƃ���ȊO�ł̕��򂪂ł��Ă��Ȃ����߁A���̕����ǋL�B
Private Function SubCurrentTestIfNeeded_f() As Double

    On Error GoTo ErrorExit
    
    '���ʕϐ�
    Dim retResult() As Double                   '2013/02/05 �C��
    Dim retResult2(nSite) As Double             '2013/02/05 �C��
    '�{���̃e�X�g���x�������B����v���d�l���ɋL�ڂ���Ă���e�X�g���x�����́A
    '�{�e�X�g�C���X�^���X��"__"�Ȃ�тɂ���ɑ�������������O���邱�Ƃœ�����B
    '   To obtain the original test label on the specification sheet. It can be
    'obtained by removing string starting with "__" from this instance name.
    Dim strResultKey As String
    strResultKey = UCase(TheExec.DataManager.InstanceName)
    
    If Go_Serial_Mesure = True Then

        '����ЂƂłЂƂ̃e�X�g����̂�SiteCheck�͕K�v
        Call SiteCheck
        
        '�ϐ���`
        Dim strPin As String                    'Arg20�@�e�X�g�[�q (Test pin name)
        Dim dblForceVoltage As Double           'Arg21�@����d�� (VDD bias value)
        Dim strSetParamCondition As String      'Arg22-1�@����p�����[�^_Opt_�����[ (�p���������茋��)
        Dim strPowerCondition As String         'Arg22-2�@PPS & Pin settings
        Dim strPatternCondition As String       'Arg22-3�@Pattern
            
        '����p�����[�^
        Dim lAve As Double                      'TestCondition
        Dim dblClampCurrent As Double           'TestCondition
        Dim dblWait As Double                   'TestCondition
        Dim dblPinResourceName As String        'TestCondition
                
        '�֐����ϐ�
        Dim Flg_Active(nSite) As Long
        Dim TempValue(nSite) As Double
        Dim site As Long
        Dim mychanType As chtype
        
        '�ϐ���荞��
        '   To obtain the argument parameters on test instances sheet.
        If Not getParam_SerialMeasureAfterParallel( _
                  strPin, _
                  dblForceVoltage, _
                  strSetParamCondition, _
                  strPowerCondition, _
                  strPatternCondition) Then
                  MsgBox "The Number of subCurrent_Test_f's arguments is invalid!"
                  Call DisableAllTest 'EeeJob�֐�
                  Exit Function
        End If
            
        '�p�����[�^�ݒ�̊֐����Ă� (FW_SetSubCurrentParam)
        '   To call measurement parameter setting condition and environment setup condition.
        Call TheCondition.SetCondition(strSetParamCondition) '1
    
        'Test Condition�Őݒ肳��Ă���p�����[�^��VarBank���擾
        '   To obtain dc measurement parameters registered in the VarBank.
        lAve = GetSubCurrentAverageCount(GetInstanceName)
        dblClampCurrent = GetSubCurrentClampCurrent(GetInstanceName)
        dblWait = GetSubCurrentWaitTime(GetInstanceName)
        dblPinResourceName = GetSubCurrentPinResourceName(GetInstanceName)
    
        'Active�T�C�g�̊m�F
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                Flg_Active(site) = 1
            End If
        Next site
        
        '�`���l���^�C�v�̊m�F
        mychanType = GetChanType(strPin)

        'SUB�d������̊m�F
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                '========== ���g�psite��BetaGND�؂藣�� ===============================
                Call GndSeparateBySite(site)
                
                '========== Set Condition ===============================
                '   To execute device setup (power and pin electronics setting)
                Call TheCondition.SetCondition(strPowerCondition) '2
    
                '========== Force Voltage ===============================
                '����BHSD200�ł����Ă�BPMU���g���܂��B
                'Measurement. Measurement will be performed with BPMU in case of digital channel,
                'regardless of HSD100 or HSD200.
                If mychanType = chIO Then
                    Call SetFVMI_BPMU(strPin, 0 * V, dblClampCurrent, site)
                    Call DisconnectPins(strPin, site)
                    Call SetFVMI_BPMU(strPin, dblForceVoltage, dblClampCurrent, site)
                Else
                    '========== Force Voltage ===============================
                    Call SetFVMI(strPin, 0 * V, dblClampCurrent, site)
                    Call DisconnectPins(strPin, site)
                    Call SetFVMI(strPin, dblForceVoltage, dblClampCurrent, site)
                End If
                
                '========== Set Pattern ===============================
                Call TheCondition.SetCondition(strPatternCondition) '3
                TheHdw.WAIT dblWait
                    
                '========== Measure Current ===============================
                If mychanType = chIO Then
                    Call MeasureI_BPMU(strPin, retResult2, lAve, site)
                    Call SetFVMI_BPMU(strPin, 0 * V, dblClampCurrent, site)
                    Call DisconnectPins(strPin, site)
                Else
                    Call MeasureI(strPin, retResult2, lAve, site)
                    Call SetFVMI(strPin, 0 * V, dblClampCurrent, site)
                    Call DisconnectPins(strPin, site)
                End If
                
                '========== ���g�psite��BetaGND�߂� =======================
                Call GndConectBySite(site, Flg_Active)
            End If
        Next site
        
      
        '�p�^�[����~
        Call StopPattern 'EeeJob�֐�
      
        'All_Open�y��Disconnect����������
        Call PowerDownAndDisconnect
                
        '�����͕Ԃ���Add����̂�; Add the result to the Result Manager.
        Call updateResult(strResultKey, retResult2)
        Call test(retResult2)
    Else
        Call TheResult.GetResult(strResultKey, retResult)
        Call test(retResult)
    End If
    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call TheResult.Add(strResultKey, retResult)

End Function


'��
'��������������������������������������
'subCurrent_Parallel��Serial_Test�p:Start
'��������������������������������������
'��
' History:  First drafted by T.Koyama 2013-12-04 (IMX219 MP1)
'           Site serial current measurement method without pattern burst.
Private Function SubCurrentTestIfNeededNoPattern_f() As Double

    On Error GoTo ErrorExit
    
    '���ʕϐ�
    Dim retResult() As Double                   '2013/02/05 �C��
    Dim retResult2(nSite) As Double             '2013/02/05 �C��
    '�{���̃e�X�g���x�������B����v���d�l���ɋL�ڂ���Ă���e�X�g���x�����́A
    '�{�e�X�g�C���X�^���X��"__"�Ȃ�тɂ���ɑ�������������O���邱�Ƃœ�����B
    '   To obtain the original test label on the specification sheet. It can be
    'obtained by removing string starting with "__" from this instance name.
    Dim strResultKey As String
    strResultKey = UCase(TheExec.DataManager.InstanceName)
    
    If Go_Serial_Mesure = True Then

        '����ЂƂłЂƂ̃e�X�g����̂�SiteCheck�͕K�v
        Call SiteCheck
        
        '�ϐ���`
        Dim strPin As String                    'Arg20�@�e�X�g�[�q (Test pin name)
        Dim dblForceVoltage As Double           'Arg21�@����d�� (VDD bias value)
        Dim strSetParamCondition As String      'Arg22-1�@����p�����[�^_Opt_�����[ (�p���������茋��)
        Dim strPowerCondition As String         'Arg22-2�@PPS & Pin settings
            
        '����p�����[�^
        Dim lAve As Double                      'TestCondition
        Dim dblClampCurrent As Double           'TestCondition
        Dim dblWait As Double                   'TestCondition
        Dim dblPinResourceName As String        'TestCondition
                
        '�֐����ϐ�
        Dim Flg_Active(nSite) As Long
        Dim TempValue(nSite) As Double
        Dim site As Long
        Dim mychanType As chtype
        
        '�ϐ���荞��
        '   To obtain the argument parameters on test instances sheet.
        If Not getParam_SerialMeasureAfterParallel_NoPattern( _
                  strPin, _
                  dblForceVoltage, _
                  strSetParamCondition, _
                  strPowerCondition) Then
                  MsgBox "The Number of subCurrent_Test_f's arguments is invalid!"
                  Call DisableAllTest 'EeeJob�֐�
                  Exit Function
        End If
            
        '�p�����[�^�ݒ�̊֐����Ă� (FW_SetSubCurrentParam)
        '   To call measurement parameter setting condition and environment setup condition.
        Call TheCondition.SetCondition(strSetParamCondition) '1
    
        'Test Condition�Őݒ肳��Ă���p�����[�^��VarBank���擾
        '   To obtain dc measurement parameters registered in the VarBank.
        lAve = GetSubCurrentAverageCount(GetInstanceName)
        dblClampCurrent = GetSubCurrentClampCurrent(GetInstanceName)
        dblWait = GetSubCurrentWaitTime(GetInstanceName)
        dblPinResourceName = GetSubCurrentPinResourceName(GetInstanceName)
    
        '�`���l���^�C�v�̊m�F
        mychanType = GetChanType(strPin)
        
        'Active�T�C�g�̊m�F
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                Flg_Active(site) = 1
            End If
        Next site
    
        'SUB�d������̊m�F
        For site = 0 To nSite
            If TheExec.sites.site(site).Active = True Then
                '========== ���g�psite��BetaGND�؂藣�� ===============================
                Call GndSeparateBySite(site)
                
                '========== Set Condition ===============================
                '   To execute device setup (power and pin electronics setting)
                Call TheCondition.SetCondition(strPowerCondition) '2

                '========== Force Voltage ===============================
                '����BHSD200�ł����Ă�BPMU���g���܂��B
                'Measurement. Measurement will be performed with BPMU in case of digital channel,
                'regardless of HSD100 or HSD200.
                'Digital Channel�Ȃ�BPMU���g�p����B�A���AHSD200�ł����Ă�BPMU���g�p����(IMX164�̑��ւ���K�v�Ɣ��f���ꂽ)
                If mychanType = chIO Then
                    Call SetFVMI_BPMU(strPin, 0 * V, dblClampCurrent, site)
                    Call DisconnectPins(strPin, site)
                    Call SetFVMI_BPMU(strPin, dblForceVoltage, dblClampCurrent, site)
                Else
                    Call SetFVMI(strPin, 0 * V, dblClampCurrent, site)
                    Call DisconnectPins(strPin, site)
                    Call SetFVMI(strPin, dblForceVoltage, dblClampCurrent, site)
                End If
                
                '========== Set Pattern ===============================
                TheHdw.WAIT dblWait
                
                '========== Measure Current ===============================
                If mychanType = chIO Then
                    Call MeasureI_BPMU(strPin, retResult2, lAve, site)
                    Call SetFVMI_BPMU(strPin, 0 * V, dblClampCurrent, site)
                    Call DisconnectPins(strPin, site)
                Else
                    Call MeasureI(strPin, retResult2, lAve, site)
                    Call SetFVMI(strPin, 0 * V, dblClampCurrent, site)
                    Call DisconnectPins(strPin, site)
                End If
    
                '========== ���g�psite��BetaGND�߂� =======================
                Call GndConectBySite(site, Flg_Active)
            End If
        Next site
     
        'All_Open�y��Disconnect����������
        Call PowerDownAndDisconnect
                
        '�����͕Ԃ���Add����̂�; Add the result to the Result Manager.
        Call updateResult(strResultKey, retResult2)
        Call test(retResult2)
    Else
        Call TheResult.GetResult(strResultKey, retResult)
        Call test(retResult)
    End If
    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call TheResult.Add(strResultKey, retResult)

End Function

Private Function updateResult(ByVal keyName As String, ByRef resultValue() As Double) As Boolean
    On Error GoTo ErrorDetected
    Call TheResult.Add(keyName, resultValue)
    updateResult = True
    Exit Function
ErrorDetected:
    Call TheResult.Delete(keyName)
    Call TheResult.Add(keyName, resultValue)
    updateResult = True
End Function

Private Function Sub_SeriParaJudge_GetParameter( _
    ByRef strResultKey As String, _
    ByRef dblLoJudgeLimit As Double, _
    ByRef dblHiJudgeLimit As Double, _
    ByRef dblHiLoLimValid As Long _
    ) As Boolean

On Error GoTo ErrHandler
    '�ϐ���荞��
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_SUB_SERIPARA_JUDGE_ARGS) Then
        Sub_SeriParaJudge_GetParameter = False
        Exit Function
    End If
    
    strResultKey = ArgArr(0)
    
    If ArgArr(1) <> "" Then
        dblLoJudgeLimit = ArgArr(1)
    Else
        ArgArr(1) = 0
    End If
    
    If ArgArr(2) <> "" Then
        dblHiJudgeLimit = ArgArr(2)
    Else
        ArgArr(2) = 0
    End If
    
    dblHiLoLimValid = ArgArr(3)
    On Error GoTo 0

    Sub_SeriParaJudge_GetParameter = True
    Exit Function
    
ErrHandler:

    Sub_SeriParaJudge_GetParameter = False
    Exit Function

End Function

Public Sub mf_Sub_SeriParaJudge(ByVal Data As String, ByRef LoLimit As Double, ByRef HiLimit As Double, ByRef HiLoLimValid As Long, ByRef Retry_flag As Boolean) '2012/11/16 175Debug Arikawa
'���@:
'   1. �p���������������
'   2. �p����������̌��ʂ�p���ĕ��ϒl�����߂�B
'
'�V���A�����p����������
'   �V���p������̓˓������l���ȉ��̕��@�Ōv�Z�����B
'   1. N�������肵�����ʂ̃q�X�g�O�����ɂ����āA���z�̍ł���������(�Ⴂ��)�Ɣ��f
'      �����l�����߂�<���z�����l>�B(���A�����ł������z�Ƃ́A�u���@�v��2�ɋL�ڂ���
'       �u���ϒl�v�̕��z�̂��ƁB
'   2. �w�肳���<�X�y�b�N�l>��p���āA�˓������l�͈ȉ��̎��œ�����B
'           <�˓������l> = <���z�����l> * <�T�C�g��> + (<�X�y�b�N�l> - <���z�����l>)
'                          ~~~~~~~~~~~~~~~~~~~~~~~~    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'       �E�ӑ�ꍀ�������ƁA���z�����l�Ȃ̂ŁA�قƂ�Ǐ�ɃV���A������ɓ˓����Ă����܂��B
'       �����Ń}�[�W���Ƃ��ėp�ӂ���Ă���̂��E�ӑ�񍀁B
'   3. ���̂悤�ɂ��Č��߂�ꂽ�˓������l�́A�p���������茋�ʂ̕��ϒl���T�C�g���{����
'       �l�ɓK�p���Ȃ���΂Ȃ�Ȃ��B�p���������茋�ʂ̍��v�l(�����Ă���T�C�g���{��������)
'      �ɓK�p����̂͂��߁B
'����
'   �����O�R���^�N�g���ɂ͕K���V���A����������{����
'
'Method:
'   1. Parallel current measurement.
'   2. Calculating the mean value of step1, judge if serial measurement is needed.
'
'Judgement
'   The limit value for the judgement is calculated based on the following idea.
'   1. Using a large amount of chip measurement result, determine a value at the lower bound
'      of measurement value histogram. (Each measurement result is the mean
'      value described at step 2 of "Method" section.
'   2. Assume that the spec value <spec>, the limit value is calculated by the equation
'             <limit value> = <lower bound value> * <nSite + 1> + (<spec> - <lower bound value>)
'                             ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'        The first term on the right only will result in execution of serial measurement
'       in almost all the cases, since <lower bound value> is at the lower bound of histogram.
'       The second term on the right, therefore, is the margin term.
'   3. The <limit value> must be applied to the measured mean value multiplied by the
'      number of sites under test (not the number of sites alive in judgement).
'Notice:
'   At least more than one site is located beyond wafer edge, serial current measurement
'   must be unconditionally executed.

    Dim site As Long
    Dim Active_site As Long
    Dim TempValue() As Double
    Dim tempValueAve(nSite) As Double
    Dim tempValueSum As Double
                
    '�p���������茋�ʂ��擾(DC Test Scenario��Operation-Result�Ɋi�[)
    'To obtain parallel measurement result (conducted by DC Test Scenario)(stored in Operation-Result).
    Call TheDcTest.GetTempResult(Data, TempValue)
'    Call TheResult.GetResult(Data, tempValue) '2012/11/16 175Debug Arikawa
    
    '�e�X�g���x���𕪉����A�{���i�[���ׂ��e�X�g���x���𒊏o����B���͂����e�X�g���x���́A
    '����v���d�l���Ɏw�肳���e�X�g���x���ɑ΂��āA"__"�Ƃ���ɑ��������񂪒ǉ�����Ă���B
    'To obtain the test label by extracting the input operation-result label name, assuming
    'that the operation-result label is made with a string starting with "__" following
    'the test label on the specification sheet.
    Dim uScorePos As Long
    Dim originalTestLabel As String
    uScorePos = InStr(Data, "__")
    originalTestLabel = Mid(Data, 1, uScorePos - 1)
    
    '========== RETRY CHECK ======================================
    Retry_flag = False

    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tempValueSum = tempValueSum + TempValue(site)
            Active_site = Active_site + 1
        End If
    Next site

    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tempValueAve(site) = tempValueSum / Active_site
        End If
    Next site

'    Call TheResult.Delete(Data) '2012/11/16 175Debug Arikawa
    '�����͕Ԃ���Add����̂�
'    Call TheResult.Add(Data, tempValueAve)  '2012/11/16 175Debug Arikawa
    Call TheResult.Add(originalTestLabel, tempValueAve)
    
'    '�V���A������˓������l�̎Z�o
'    tempValueSum = tempValueSum / Active_site * SITE_MAX
    
    Select Case HiLoLimValid
        Case 0
            If Active_site < SITE_MAX Then
                Retry_flag = True
            End If
        Case 1
            If tempValueSum < LoLimit Or Active_site < SITE_MAX Then
                Retry_flag = True
            End If
        Case 2
            If tempValueSum > HiLimit Or Active_site < SITE_MAX Then
                Retry_flag = True
            End If
        Case 3
            If (tempValueSum < LoLimit Or tempValueSum > HiLimit) Or Active_site < SITE_MAX Then
                Retry_flag = True
            End If
        Case Else
    End Select
    
End Sub

'��
'��������������������������������������
'subCurrent_Parallel��Serial_Test�p:End
'��������������������������������������
'��
'��
'��������������������������������������
'subCurrent_Serial_NoPattern_Test�p:Start
'��������������������������������������
'2013/02/07
'��
Private Function subCurrent_Serial_NoPattern_Test_f() As Double

    On Error GoTo ErrorExit

    '����ЂƂłЂƂ̃e�X�g����̂�SiteCheck�͕K�v
    Call SiteCheck
        
    '�ϐ���`
    Dim strResultKey As String              'Arg20�@���ږ�
    Dim strPin As String                    'Arg21�@�e�X�g�[�q
    Dim dblForceVoltage As Double           'Arg22�@����d��
    Dim strSetParamCondition As String      'Arg23�@����p�����[�^_Opt_�����[
    Dim strPowerCondition As String         'Arg24�@Set_Voltage_�[�q�ݒ�
'    Dim strPatternCondition As String       'Arg25�@Pattern
            
    '����p�����[�^
    Dim lAve As Double                      'TestCondition
    Dim dblClampCurrent As Double           'TestCondition
    Dim dblWait As Double                   'TestCondition
    Dim dblPinResourceName As String        'TestCondition
        
    '���ʕϐ�
    Dim retResult(nSite) As Double
            
    '�֐����ϐ�
    Dim Flg_Active(nSite) As Long
    Dim TempValue(nSite) As Double
    Dim site As Long
            
    '�ϐ���荞��
    If Not SubCurrentTest_NoPattern_GetParameter( _
                strResultKey, _
                strPin, _
                dblForceVoltage, _
                strSetParamCondition, _
                strPowerCondition) Then
                MsgBox "The Number of subCurrent_Serial_NoPattern_Test_f's arguments is invalid!"
                Call DisableAllTest 'EeeJob�֐�
                Exit Function
    End If
            
    '�p�����[�^�ݒ�̊֐����Ă� (FW_SetSubCurrentParam)
    Call TheCondition.SetCondition(strSetParamCondition)

    lAve = GetSubCurrentAverageCount(GetInstanceName)
    dblClampCurrent = GetSubCurrentClampCurrent(GetInstanceName)
    dblWait = GetSubCurrentWaitTime(GetInstanceName)
    dblPinResourceName = GetSubCurrentPinResourceName(GetInstanceName)

    'Active�T�C�g�̊m�F
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            Flg_Active(site) = 1
        End If
    Next site

    'SUB�d������̊m�F
    For site = 0 To nSite
        If TheExec.sites.site(site).Active = True Then
            '========== ���g�psite��BetaGND�؂藣�� ===============================
            Call GndSeparateBySite(site)
            
            '========== Set Condition ===============================
            Call TheCondition.SetCondition(strPowerCondition)

             '========== Force Voltage ===============================
            If StrComp(dblPinResourceName, "BPMU", vbTextCompare) = 0 Then
                Call SetFVMI_BPMU(strPin, 0 * V, dblClampCurrent, site)
                Call DisconnectPins(strPin, site)
                Call SetFVMI_BPMU(strPin, dblForceVoltage, dblClampCurrent, site)
            Else
                Call SetFVMI(strPin, 0 * V, dblClampCurrent, site)
                Call DisconnectPins(strPin, site)
                Call SetFVMI(strPin, dblForceVoltage, dblClampCurrent, site)
            End If
            
            TheHdw.WAIT dblWait
            
             '========== Measure Current ===============================
            If StrComp(dblPinResourceName, "BPMU", vbTextCompare) = 0 Then
                Call MeasureI_BPMU(strPin, retResult, lAve, site)
                Call SetFVMI_BPMU(strPin, 0 * V, dblClampCurrent, site)
                Call DisconnectPins(strPin, site)
            Else
                Call MeasureI(strPin, retResult, lAve, site)
                Call SetFVMI(strPin, 0 * V, dblClampCurrent, site)
                Call DisconnectPins(strPin, site)
            End If

             '========== ���g�psite��BetaGND�߂� =======================
            Call GndConectBySite(site, Flg_Active)
        End If
    Next site
    
    '�p�^�[����~
    Call StopPattern 'EeeJob�֐�
  
    'All_Open�y��Disconnect����������
    Call PowerDownAndDisconnect
           
    '�������Ԃ�Add������
    Call test(retResult)
    Call TheResult.Add(strResultKey, retResult)

    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call test(retResult)
    Call TheResult.Add(strResultKey, retResult)

End Function
'��
'��������������������������������������
'SubCurrentTest_NoPattern_GetParameter�p
'��������������������������������������
'��
'2013/02/06
Private Function SubCurrentTest_NoPattern_GetParameter( _
    ByRef strResultKey As String, _
    ByRef strPin As String, _
    ByRef dblForceVoltage As Double, _
    ByRef strSetParamCondition As String, _
    ByRef strPowerCondition As String _
    ) As Boolean

    '�ϐ���荞��
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_SUBCURRENT_ARGS) Then
        SubCurrentTest_NoPattern_GetParameter = False
        Exit Function
    End If
    
    Dim tempstr As String      '�ꎞ�ۑ��ϐ�
    Dim tempArrstr() As String '�ꎞ�ۑ��z��
    
On Error GoTo ErrHandler
    strResultKey = ArgArr(0)                'Arg20: Test label name.
    strPin = ArgArr(1)                      'Arg21: Test pin name
    dblForceVoltage = CDbl(ArgArr(2))       'Arg22: Force voltage (PPS value) for test pin.
    strSetParamCondition = ArgArr(3)        'Arg23: [Test Condition]'s condition name for common environment setup.
    strPowerCondition = ArgArr(4)           'Arg24: [Test Condition]'s condition name for device setup.
On Error GoTo 0

    SubCurrentTest_NoPattern_GetParameter = True
    Exit Function
    
ErrHandler:

    SubCurrentTest_NoPattern_GetParameter = False
    Exit Function

End Function


'��
'��������������������������������������
'subCurrent_Parallel��BPMU Serial_Judge�p:Start
'��������������������������������������
'��

Private Function subCurrentNonScenarioSeriParaJudge_f() As Double

    '����ЂƂłЂƂ̃e�X�g����̂�SiteCheck�͕K�v
    Call SiteCheck
    
    Go_Serial_Mesure = False
    
    '�ϐ���`
    Dim strResultKey As String              'Arg20�@Test label
    Dim dblLoJudgeLimit As Double           'Arg21�@Serial�˓�Low���~�b�g
    Dim dblHiJudgeLimit As Double           'Arg22�@Serial�˓�High���~�b�g
    Dim dblHiLoLimValid As Long             'Arg23�@Serial�˓����~�b�g�̗L���͈�
    

    Dim temp_strResultKey As String
    temp_strResultKey = UCase(TheExec.DataManager.InstanceName)
    Dim uScorePos As Long
    Dim originalTestLabel As String
    uScorePos = InStr(temp_strResultKey, "__")
    originalTestLabel = Mid(temp_strResultKey, 1, uScorePos - 1)
    strResultKey = originalTestLabel & "__SeriParaParaMeasure"
    
    
    '�ϐ���荞��
    If Not Sub_NonScenarioParaJudge_GetParameter( _
                dblLoJudgeLimit, _
                dblHiJudgeLimit, _
                dblHiLoLimValid) Then
                MsgBox "The Number of subCurrentSeriParaJudge_f's arguments is invalid!"
                Call DisableAllTest 'EeeJob�֐�
                Exit Function
    End If
    
    'SeriParaJudge
    Call mf_Sub_SeriParaBPMUJudge(strResultKey, dblLoJudgeLimit, dblHiJudgeLimit, dblHiLoLimValid, Go_Serial_Mesure)
    
End Function

Public Sub mf_Sub_SeriParaBPMUJudge(ByVal Data As String, ByRef LoLimit As Double, ByRef HiLimit As Double, ByRef HiLoLimValid As Long, ByRef Retry_flag As Boolean) '2012/11/16 175Debug Arikawa
'���@:
'   1. �p���������������
'   2. �p����������̌��ʂ�p���ĕ��ϒl�����߂�B
'
'�V���A�����p����������
'   �V���p������̓˓������l���ȉ��̕��@�Ōv�Z�����B
'   1. N�������肵�����ʂ̃q�X�g�O�����ɂ����āA���z�̍ł���������(�Ⴂ��)�Ɣ��f
'      �����l�����߂�<���z�����l>�B(���A�����ł������z�Ƃ́A�u���@�v��2�ɋL�ڂ���
'       �u���ϒl�v�̕��z�̂��ƁB
'   2. �w�肳���<�X�y�b�N�l>��p���āA�˓������l�͈ȉ��̎��œ�����B
'           <�˓������l> = <���z�����l> * <�T�C�g��> + (<�X�y�b�N�l> - <���z�����l>)
'                          ~~~~~~~~~~~~~~~~~~~~~~~~    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'       �E�ӑ�ꍀ�������ƁA���z�����l�Ȃ̂ŁA�قƂ�Ǐ�ɃV���A������ɓ˓����Ă����܂��B
'       �����Ń}�[�W���Ƃ��ėp�ӂ���Ă���̂��E�ӑ�񍀁B
'   3. ���̂悤�ɂ��Č��߂�ꂽ�˓������l�́A�p���������茋�ʂ̕��ϒl���T�C�g���{����
'       �l�ɓK�p���Ȃ���΂Ȃ�Ȃ��B�p���������茋�ʂ̍��v�l(�����Ă���T�C�g���{��������)
'      �ɓK�p����̂͂��߁B
'����
'   �����O�R���^�N�g���ɂ͕K���V���A����������{����
'
'Method:
'   1. Parallel current measurement.
'   2. Calculating the mean value of step1, judge if serial measurement is needed.
'
'Judgement
'   The limit value for the judgement is calculated based on the following idea.
'   1. Using a large amount of chip measurement result, determine a value at the lower bound
'      of measurement value histogram. (Each measurement result is the mean
'      value described at step 2 of "Method" section.
'   2. Assume that the spec value <spec>, the limit value is calculated by the equation
'             <limit value> = <lower bound value> * <nSite + 1> + (<spec> - <lower bound value>)
'                             ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'        The first term on the right only will result in execution of serial measurement
'       in almost all the cases, since <lower bound value> is at the lower bound of histogram.
'       The second term on the right, therefore, is the margin term.
'   3. The <limit value> must be applied to the measured mean value multiplied by the
'      number of sites under test (not the number of sites alive in judgement).
'Notice:
'   At least more than one site is located beyond wafer edge, serial current measurement
'   must be unconditionally executed.

    Dim site As Long
    Dim Active_site As Long
    Dim TempValue() As Double
    Dim tempValueAve(nSite) As Double
    Dim tempValueSum As Double
                
    '�p���������茋�ʂ��擾(DC Test Scenario��Operation-Result�Ɋi�[)
    'To obtain parallel measurement result (conducted by DC Test Scenario)(stored in Operation-Result).
'    Call TheDcTest.GetTempResult(Data, tempValue)
    Call TheResult.GetResult(Data, TempValue) '2012/11/16 175Debug Arikawa
    
    '�e�X�g���x���𕪉����A�{���i�[���ׂ��e�X�g���x���𒊏o����B���͂����e�X�g���x���́A
    '����v���d�l���Ɏw�肳���e�X�g���x���ɑ΂��āA"__"�Ƃ���ɑ��������񂪒ǉ�����Ă���B
    'To obtain the test label by extracting the input operation-result label name, assuming
    'that the operation-result label is made with a string starting with "__" following
    'the test label on the specification sheet.
    Dim uScorePos As Long
    Dim originalTestLabel As String
    uScorePos = InStr(Data, "__")
    originalTestLabel = Mid(Data, 1, uScorePos - 1)
    
    '========== RETRY CHECK ======================================
    Retry_flag = False
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tempValueSum = tempValueSum + TempValue(site)
            Active_site = Active_site + 1
        End If
    Next site

    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            tempValueAve(site) = tempValueSum / Active_site
        End If
    Next site

'    Call TheResult.Delete(Data) '2012/11/16 175Debug Arikawa
    '�����͕Ԃ���Add����̂�
'    Call TheResult.Add(Data, tempValueAve)  '2012/11/16 175Debug Arikawa
    Call TheResult.Add(originalTestLabel, tempValueAve)
    
'    '�V���A������˓������l�̎Z�o
'    tempValueSum = tempValueSum / Active_site * SITE_MAX
    
    Select Case HiLoLimValid
        Case 0
            If Active_site < SITE_MAX Then
                Retry_flag = True
            End If
        Case 1
            If tempValueSum < LoLimit Or Active_site < SITE_MAX Then
                Retry_flag = True
            End If
        Case 2
            If tempValueSum > HiLimit Or Active_site < SITE_MAX Then
                Retry_flag = True
            End If
        Case 3
            If (tempValueSum < LoLimit Or tempValueSum > HiLimit) Or Active_site < SITE_MAX Then
                Retry_flag = True
            End If
        Case Else
    End Select
    
End Sub

'��
'��������������������������������������
'subCurrent_Paralle  BPMU 2013/3/7  Hamada
'��������������������������������������
'��

Private Function SubCurrentNonScenario_Measure_f() As Double

    On Error GoTo ErrorExit
    
    '���ʕϐ�
    Dim retResult() As Double                   '2013/02/05 �C��
    Dim retResult2(nSite) As Double             '2013/02/05 �C��
    '�{���̃e�X�g���x�������B����v���d�l���ɋL�ڂ���Ă���e�X�g���x�����́A
    '�{�e�X�g�C���X�^���X��"__"�Ȃ�тɂ���ɑ�������������O���邱�Ƃœ�����B
    '   To obtain the original test label on the specification sheet. It can be
    'obtained by removing string starting with "__" from this instance name.
    Dim temp_strResultKey As String
    Dim strResultKey As String
    temp_strResultKey = UCase(TheExec.DataManager.InstanceName)
'    strResultKey = UCase(TheExec.DataManager.instanceName)
    Dim uScorePos As Long
    Dim originalTestLabel As String
    uScorePos = InStr(temp_strResultKey, "__")
    originalTestLabel = Mid(temp_strResultKey, 1, uScorePos - 1)
    strResultKey = originalTestLabel & "__SeriParaParaMeasure"

        '����ЂƂłЂƂ̃e�X�g����̂�SiteCheck�͕K�v
    Call SiteCheck
    
    '�ϐ���`
    Dim strPin As String                    'Arg20�@�e�X�g�[�q (Test pin name)
    Dim dblForceVoltage As Double           'Arg21�@����d�� (VDD bias value)
    Dim strSetParamCondition As String      'Arg22-1�@����p�����[�^_Opt_�����[ (�p���������茋��)
    Dim strPowerCondition As String         'Arg22-2�@PPS & Pin settings
    Dim strPatternCondition As String       'Arg22-3�@Pattern
        
    '����p�����[�^
    Dim lAve As Double                      'TestCondition
    Dim dblClampCurrent As Double           'TestCondition
    Dim dblWait As Double                   'TestCondition
    Dim dblPinResourceName As String        'TestCondition
            
    '�֐����ϐ�
    Dim Flg_Active(nSite) As Long
'    Dim tempValue(nSite) As Double
'    Dim strResultKey As String
    Dim site As Long
        
    Dim mychanType As chtype
    Dim FunctionName As String
    '�ϐ���荞��
    '   To obtain the argument parameters on test instances sheet.
    If Not getParam_BPMU_Parallel( _
              strPin, _
              dblForceVoltage, _
              strSetParamCondition, _
              strPowerCondition, _
              strPatternCondition _
              ) Then
              MsgBox "The Number of subCurrent_Test_f's arguments is invalid!"
              Call DisableAllTest 'EeeJob�֐�
              Exit Function
    End If
    
    Dim Temp_ChanArr() As Long
    Dim Temp_chanCnt As Long
    Dim Temp_BoardCnt As Long
    Dim Temp_ChanbyBoard As Long
    Dim Temp_Message As String

    Dim tempValueSum As Double
    Dim TempValue(nSite) As Double
    
    Dim i As Long
    Dim ii As Long
    Dim myboard_i As Long
    myboard_i = 0
    
    
    Dim Active_site As Long
    Active_site = 0
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            Active_site = Active_site + 1
        End If
    Next site
    
    
    If TesterType <> "IP750EX" Then
        
        mychanType = GetChanType(strPin)
        Call TheExec.DataManager.GetChanListByBoard(strPin, ALL_SITE, mychanType, Temp_ChanArr, Temp_chanCnt, Temp_BoardCnt, Temp_ChanbyBoard, Temp_Message)
        
        Call TheCondition.SetCondition(strSetParamCondition) '1
    
        'Test Condition�Őݒ肳��Ă���p�����[�^��VarBank���擾
        lAve = GetSubCurrentAverageCount(GetInstanceName)
        dblClampCurrent = GetSubCurrentClampCurrent(GetInstanceName)
        dblWait = GetSubCurrentWaitTime(GetInstanceName)
        dblPinResourceName = GetSubCurrentPinResourceName(GetInstanceName)
        
             If mychanType = chIO Then
'                Dim cBoard As Long
'                For cBoard = 0 To Temp_BoardCnt - 1

                '========== Set Condition ===============================
                Call TheCondition.SetCondition(strPowerCondition) '2
        
                '========== Force Voltage ===============================
                Call SetFVMI_BPMU(strPin, dblForceVoltage, dblClampCurrent)
        
                '========== Set Pattern ===============================
                Call TheCondition.SetCondition(strPatternCondition) '3
                TheHdw.WAIT dblWait
                
                '========== Measure Current ===============================
                Call MeasureI_BPMU(strPin, TempValue, lAve)
                                
                For i = 0 To Temp_BoardCnt - 1
                        tempValueSum = tempValueSum + TempValue(myboard_i)
                        myboard_i = myboard_i + Temp_ChanbyBoard
                Next i
                
                For site = 0 To nSite
                    If TheExec.sites.site(site).Active = True Then
                            retResult2(site) = tempValueSum
                    End If
                Next site
                
            Else
                '========== Set Condition ===============================
                Call TheCondition.SetCondition(strPowerCondition) '2
                '========== Force Voltage ===============================
                Call SetFVMI(strPin, dblForceVoltage, dblClampCurrent)
                '========== Set Pattern ===============================
                Call TheCondition.SetCondition(strPatternCondition) '3
                TheHdw.WAIT dblWait
                '========== Measure Current ===============================
                Call MeasureI(strPin, retResult2, lAve)
            End If
         
            
        '�p�^�[����~
        Call StopPattern 'EeeJob�֐�
        
        Call DisconnectPins(strPin, ALL_SITE)
        'All_Open�y��Disconnect����������
        Call PowerDownAndDisconnect
        
    Else
    
    
        Call TheCondition.SetCondition(strSetParamCondition) '1
        lAve = GetSubCurrentAverageCount(GetInstanceName)
        dblClampCurrent = GetSubCurrentClampCurrent(GetInstanceName)
        dblWait = GetSubCurrentWaitTime(GetInstanceName)
        dblPinResourceName = GetSubCurrentPinResourceName(GetInstanceName)
        
        '========== Set Condition ===============================
        Call TheCondition.SetCondition(strPowerCondition) '2
        '========== Force Voltage ===============================
        Call SetFVMI(strPin, dblForceVoltage, dblClampCurrent)
        '========== Set Pattern ===============================
        Call TheCondition.SetCondition(strPatternCondition) '3
        TheHdw.WAIT dblWait
        '========== Measure Current ===============================
        Call MeasureI(strPin, retResult2, lAve, site)
            
        '�p�^�[����~
        Call StopPattern 'EeeJob�֐�
        Call DisconnectPins(strPin, ALL_SITE)
        Call PowerDownAndDisconnect
        
    End If

    Call updateResult(strResultKey, retResult2)

    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call TheResult.Add(strResultKey, retResult)

End Function
'��
'��������������������������������������
'subCurrent_Paralle  BPMU 2013/3/7  Hamada
'��������������������������������������
'��

Private Function SubCurrentNonScenarioNoPattern_Measure_f() As Double

    On Error GoTo ErrorExit
    
    '���ʕϐ�
    Dim retResult() As Double                   '2013/02/05 �C��
    Dim retResult2(nSite) As Double             '2013/02/05 �C��
    '�{���̃e�X�g���x�������B����v���d�l���ɋL�ڂ���Ă���e�X�g���x�����́A
    '�{�e�X�g�C���X�^���X��"__"�Ȃ�тɂ���ɑ�������������O���邱�Ƃœ�����B
    '   To obtain the original test label on the specification sheet. It can be
    'obtained by removing string starting with "__" from this instance name.
    Dim temp_strResultKey As String
    Dim strResultKey As String
    temp_strResultKey = UCase(TheExec.DataManager.InstanceName)
'    strResultKey = UCase(TheExec.DataManager.instanceName)
    Dim uScorePos As Long
    Dim originalTestLabel As String
    uScorePos = InStr(temp_strResultKey, "__")
    originalTestLabel = Mid(temp_strResultKey, 1, uScorePos - 1)
    strResultKey = originalTestLabel & "__SeriParaParaMeasure"

        '����ЂƂłЂƂ̃e�X�g����̂�SiteCheck�͕K�v
    Call SiteCheck
    
    '�ϐ���`
    Dim strPin As String                    'Arg20�@�e�X�g�[�q (Test pin name)
    Dim dblForceVoltage As Double           'Arg21�@����d�� (VDD bias value)
    Dim strSetParamCondition As String      'Arg22-1�@����p�����[�^_Opt_�����[ (�p���������茋��)
    Dim strPowerCondition As String         'Arg22-2�@PPS & Pin settings
    Dim strPatternCondition As String       'Arg22-3�@Pattern
        
    '����p�����[�^
    Dim lAve As Double                      'TestCondition
    Dim dblClampCurrent As Double           'TestCondition
    Dim dblWait As Double                   'TestCondition
    Dim dblPinResourceName As String        'TestCondition
            
    '�֐����ϐ�
    Dim Flg_Active(nSite) As Long
'    Dim tempValue(nSite) As Double
'    Dim strResultKey As String
    Dim site As Long
        
    Dim mychanType As chtype
    Dim FunctionName As String
    '�ϐ���荞��
    '   To obtain the argument parameters on test instances sheet.
    If Not getParam_BPMU_Parallel_NonPattern( _
              strPin, _
              dblForceVoltage, _
              strSetParamCondition, _
              strPowerCondition _
              ) Then
              MsgBox "The Number of subCurrent_Test_f's arguments is invalid!"
              Call DisableAllTest 'EeeJob�֐�
              Exit Function
    End If
    
    Dim Temp_ChanArr() As Long
    Dim Temp_chanCnt As Long
    Dim Temp_BoardCnt As Long
    Dim Temp_ChanbyBoard As Long
    Dim Temp_Message As String

    Dim tempValueSum As Double
    Dim TempValue(nSite) As Double
    
    Dim i As Long
    Dim ii As Long
    Dim myboard_i As Long
    myboard_i = 0
    
    
    Dim Active_site As Long
    Active_site = 0
    
    For site = 0 To nSite
        If TheExec.sites.site(site).Active Then
            Active_site = Active_site + 1
        End If
    Next site
    
    
    If TesterType <> "IP750EX" Then
        
        mychanType = GetChanType(strPin)
        Call TheExec.DataManager.GetChanListByBoard(strPin, ALL_SITE, mychanType, Temp_ChanArr, Temp_chanCnt, Temp_BoardCnt, Temp_ChanbyBoard, Temp_Message)
        
        Call TheCondition.SetCondition(strSetParamCondition) '1
    
        'Test Condition�Őݒ肳��Ă���p�����[�^��VarBank���擾
        lAve = GetSubCurrentAverageCount(GetInstanceName)
        dblClampCurrent = GetSubCurrentClampCurrent(GetInstanceName)
        dblWait = GetSubCurrentWaitTime(GetInstanceName)
        dblPinResourceName = GetSubCurrentPinResourceName(GetInstanceName)
        
             If mychanType = chIO Then
'                Dim cBoard As Long
'                For cBoard = 0 To Temp_BoardCnt - 1

                '========== Set Condition ===============================
                Call TheCondition.SetCondition(strPowerCondition) '2
        
                '========== Force Voltage ===============================
                Call SetFVMI_BPMU(strPin, dblForceVoltage, dblClampCurrent)
        
                '========== Set Pattern ===============================
                TheHdw.WAIT dblWait
                
                '========== Measure Current ===============================
                Call MeasureI_BPMU(strPin, TempValue, lAve)
                                
                For i = 0 To Temp_BoardCnt - 1
                        tempValueSum = tempValueSum + TempValue(myboard_i)
                        myboard_i = myboard_i + Temp_ChanbyBoard
                Next i
                
                For site = 0 To nSite
                    If TheExec.sites.site(site).Active = True Then
                            retResult2(site) = tempValueSum
                    End If
                Next site
                
            Else
                '========== Set Condition ===============================
                Call TheCondition.SetCondition(strPowerCondition) '2
                '========== Force Voltage ===============================
                Call SetFVMI(strPin, dblForceVoltage, dblClampCurrent)
                '========== Set Pattern ===============================
                TheHdw.WAIT dblWait
                '========== Measure Current ===============================
                Call MeasureI(strPin, retResult2, lAve)
            End If
         
            
        '�p�^�[����~
        
        Call DisconnectPins(strPin, ALL_SITE)
        'All_Open�y��Disconnect����������
        Call PowerDownAndDisconnect
        
    Else
    
    
        Call TheCondition.SetCondition(strSetParamCondition) '1
        lAve = GetSubCurrentAverageCount(GetInstanceName)
        dblClampCurrent = GetSubCurrentClampCurrent(GetInstanceName)
        dblWait = GetSubCurrentWaitTime(GetInstanceName)
        dblPinResourceName = GetSubCurrentPinResourceName(GetInstanceName)
        
        '========== Set Condition ===============================
        Call TheCondition.SetCondition(strPowerCondition) '2
        '========== Force Voltage ===============================
        Call SetFVMI(strPin, dblForceVoltage, dblClampCurrent)
        '========== Set Pattern ===============================
        TheHdw.WAIT dblWait
        '========== Measure Current ===============================
        Call MeasureI(strPin, retResult2, lAve)
            
        '�p�^�[����~
        Call DisconnectPins(strPin, ALL_SITE)
        Call PowerDownAndDisconnect
        
    End If

    Call updateResult(strResultKey, retResult2)

    Exit Function
            
ErrorExit:
    MsgBox "Error Occured !! " & CStr(Err.Number) & " - " & Err.Source & Chr(13) & Chr(13) & Err.Description
    Call DisableAllTest 'EeeJob�֐�
    Call TheResult.Add(strResultKey, retResult)

End Function


Private Function GetChanType(ByVal PinList As String) As chtype

    Dim chanType As chtype
    Dim pinArr() As String
    Dim pinNum As Long
    Dim i As Long

    With TheExec.DataManager
        chanType = .chanType(PinList)

        If chanType = chUnk Then
            On Error GoTo INVALID_PINLIST
            Call .DecomposePinList(PinList, pinArr, pinNum)
            On Error GoTo -1

            chanType = .chanType(pinArr(0))
            For i = 0 To pinNum - 1
                If chanType <> .chanType(pinArr(i)) Then
                    chanType = chUnk
                    Exit For
                End If
            Next i
        End If
    End With

    GetChanType = chanType
    Exit Function

INVALID_PINLIST:
    GetChanType = chUnk

End Function


'��
'��������������������������������������
'subCurrent_Serial_Test�p:End
'��������������������������������������
'��
Private Function getParam_BPMU_Parallel( _
    ByRef strPin As String, _
    ByRef dblForceVoltage As Double, _
    ByRef strSetParamCondition As String, _
    ByRef strPowerCondition As String, _
    ByRef strPatternCondition As String _
    ) As Boolean

    '�ϐ���荞��
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_BPMU_PARA_ARGS) Then
        getParam_BPMU_Parallel = False
        Exit Function
    End If
    
    Dim tempstr As String      '�ꎞ�ۑ��ϐ�
    Dim tempArrstr() As String '�ꎞ�ۑ��z��
    
On Error GoTo ErrHandler
    strPin = ArgArr(0)                      'Arg20: Test pin name
    dblForceVoltage = CDbl(ArgArr(1))       'Arg21: Force voltage (PPS value) for test pin.
    tempstr = ArgArr(2)
    tempArrstr = Split(tempstr, ",")
    strSetParamCondition = tempArrstr(0)        'Arg22-1: [Test Condition]'s condition name for common environment setup.
    strPowerCondition = tempArrstr(1)           'Arg22-2: [Test Condition]'s condition name for device setup.
    strPatternCondition = tempArrstr(2)         'Arg22-3: [Test Condition]'s condition name for patttern burst.
'    strResultKey = ArgArr(3)
On Error GoTo 0

    getParam_BPMU_Parallel = True
    Exit Function
    
ErrHandler:

    getParam_BPMU_Parallel = False
    Exit Function

End Function



'��
'��������������������������������������
'subCurrent_Serial_Test�p:End
'��������������������������������������
'��
Private Function getParam_BPMU_Parallel_NonPattern( _
    ByRef strPin As String, _
    ByRef dblForceVoltage As Double, _
    ByRef strSetParamCondition As String, _
    ByRef strPowerCondition As String _
    ) As Boolean

    '�ϐ���荞��
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, EEE_AUTO_BPMU_PARA_ARGS) Then
        getParam_BPMU_Parallel_NonPattern = False
        Exit Function
    End If
    
    Dim tempstr As String      '�ꎞ�ۑ��ϐ�
    Dim tempArrstr() As String '�ꎞ�ۑ��z��
    
On Error GoTo ErrHandler
    strPin = ArgArr(0)                      'Arg20: Test pin name
    dblForceVoltage = CDbl(ArgArr(1))       'Arg21: Force voltage (PPS value) for test pin.
    tempstr = ArgArr(2)
    tempArrstr = Split(tempstr, ",")
    strSetParamCondition = tempArrstr(0)        'Arg22-1: [Test Condition]'s condition name for common environment setup.
    strPowerCondition = tempArrstr(1)           'Arg22-2: [Test Condition]'s condition name for device setup.
'    strResultKey = ArgArr(3)
On Error GoTo 0

    getParam_BPMU_Parallel_NonPattern = True
    Exit Function
    
ErrHandler:

    getParam_BPMU_Parallel_NonPattern = False
    Exit Function

End Function

Private Function Sub_NonScenarioParaJudge_GetParameter( _
    ByRef dblLoJudgeLimit As Double, _
    ByRef dblHiJudgeLimit As Double, _
    ByRef dblHiLoLimValid As Long _
    ) As Boolean

On Error GoTo ErrHandler
    '�ϐ���荞��
    '�z�萔���ƈႤ�ꍇ�G���[�R�[�h
    Dim ArgArr() As String
    If Not EeeAutoGetArgument(ArgArr, 4) Then
        Sub_NonScenarioParaJudge_GetParameter = False
        Exit Function
    End If
    
'    strResultKey = ArgArr(0)
    
    If ArgArr(0) <> "" Then
        dblLoJudgeLimit = ArgArr(0)
    Else
        ArgArr(0) = 0
    End If
    
    If ArgArr(1) <> "" Then
        dblHiJudgeLimit = ArgArr(1)
    Else
        ArgArr(1) = 0
    End If
    
    dblHiLoLimValid = ArgArr(2)
    On Error GoTo 0

    Sub_NonScenarioParaJudge_GetParameter = True
    Exit Function
    
ErrHandler:

    Sub_NonScenarioParaJudge_GetParameter = False
    Exit Function

End Function
